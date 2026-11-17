import io
import math
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import ezdxf
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from ezdxf import recover
from ezdxf.addons import odafc


# ==============================================================================
# 1. UU CAD SCANNER ENGINE (2D CROSS-SECTION)
# ==============================================================================

@dataclass
class UUScanConfig:
    unit_to_meters: float = 0.001
    vertical_exaggeration: float = 1.0
    datum_cad_y: float = 0.0
    datum_elevation_m: float = 0.0
    leader_snap_radius_m: float = 0.50
    text_search_radius_m: float = 2.50
    min_clearance_m: float = 0.30
    ground_layer_keywords: tuple[str, ...] = ("GL", "GROUND", "EXIST_GROUND", "TOPO")


DIAMETER_RE = re.compile(r"(?:DN|DIA\.?|Ø|\b)\s*(\d{2,4})\s*(?:MM|DIA)?\b", re.IGNORECASE)
MATERIAL_RE = re.compile(r"\b(DI|HDPE|UPVC|PVC|RC|GI|VCP|PE|STEEL|CI)\b", re.IGNORECASE)
INVERT_RE = re.compile(r"\bIL\s*[:=]?\s*([+-]?\d+(?:\.\d+)?)", re.IGNORECASE)
GROUND_RE = re.compile(r"\bGL\s*[:=]?\s*([+-]?\d+(?:\.\d+)?)", re.IGNORECASE)

UTILITY_KEYWORDS: dict[str, tuple[str, ...]] = {
    "Fresh Water": ("FW", "FRESH", "WSD-FW", "POTABLE"),
    "Salt Water": ("SW", "SALT", "FLUSH", "WSD-SW"),
    "Foul Sewer": ("FOUL", "FS", "SEW", "DSD-F"),
    "Stormwater": ("STORM", "STW", "SWD", "DRAIN", "DSD-S"),
    "Gas": ("GAS", "HKCG", "TOWNGAS"),
    "Power": ("HV", "LV", "11KV", "132KV", "ELEC", "CLP", "HEC"),
    "Telecom": ("TEL", "COMMS", "FIBER", "PCCW", "HGC"),
}

UTILITY_COLORS: dict[str, str] = {
    "Fresh Water": "#1f77b4",
    "Salt Water": "#17becf",
    "Foul Sewer": "#8c564b",
    "Stormwater": "#2ca02c",
    "Gas": "#ffbf00",
    "Power": "#d62728",
    "Telecom": "#9467bd",
    "Unclassified": "#7f7f7f",
}


def generate_sample_uu_dxf_bytes() -> bytes:
    """Create a realistic sample UU cross-section DXF in memory (units: mm)."""
    doc = ezdxf.new("R2018")
    msp = doc.modelspace()

    # Ground surface polyline at +5.00 mPD (Y = 5000 mm)
    msp.add_lwpolyline([(-4000, 5000), (0, 5020), (4000, 4980)], dxfattribs={"layer": "EXIST_GROUND"})

    # Existing UU pipes: (x_mm, center_y_mm, radius_mm, layer, callout_text)
    sample_pipes = [
        (-1800, 3900, 150, "WSD-FW", "DN300 DI FW (IL +3.75mPD)"),
        (-650, 4150, 75, "CLP-HV", "150mm DUCT 11kV HV POWER"),
        (350, 3850, 100, "HKCG-GAS", "DN200 PE GAS MAIN"),
        (1650, 3300, 225, "DSD-STORM", "DN450 RC STW DRAIN"),
        (2600, 4200, 55, "PCCW-TEL", "DN110 uPVC TEL DUCT"),
    ]

    for x, y, r, layer, label in sample_pipes:
        msp.add_circle((x, y), r, dxfattribs={"layer": layer})
        # Add a leader pointing from pipe crown up-right to callout text
        tip = (x + r * 0.7, y + r * 0.7)
        elbow = (x + 450, y + 450)
        msp.add_leader([tip, elbow], dxfattribs={"layer": f"{layer}-TEXT"})
        msp.add_text(
            label,
            dxfattribs={"layer": f"{layer}-TEXT", "height": 110, "insert": (elbow[0] + 40, elbow[1])},
        )

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


def load_cad_from_bytes(file_bytes: bytes, filename: str) -> ezdxf.document.Drawing:
    """Load .dxf or .dwg from uploaded file bytes via a temporary file."""
    suffix = Path(filename).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file_bytes)
        tmp_path = Path(tmp.name)

    try:
        if suffix == ".dwg":
            if not odafc.is_installed():
                raise RuntimeError("ODA File Converter not installed on server. Please upload a .dxf file.")
            return odafc.readfile(str(tmp_path))
        if suffix == ".dxf":
            try:
                return ezdxf.readfile(str(tmp_path))
            except ezdxf.DXFStructureError:
                doc, _ = recover.readfile(str(tmp_path))
                return doc
        raise ValueError(f"Unsupported file type: {suffix}")
    finally:
        tmp_path.unlink(missing_ok=True)


def classify_utility(layer: str, text: str) -> str:
    combined = f"{layer} {text}".upper()
    for utility_name, tokens in UTILITY_KEYWORDS.items():
        if any(re.search(rf"(?:^|[^A-Z0-9]){re.escape(tok)}(?:$|[^A-Z0-9])", combined) for tok in tokens):
            return utility_name
    return "Unclassified"


def parse_callout_specs(text: str) -> dict[str, Any]:
    dia_match = DIAMETER_RE.search(text)
    mat_match = MATERIAL_RE.search(text)
    il_match = INVERT_RE.search(text)
    gl_match = GROUND_RE.search(text)
    return {
        "annotated_diameter_mm": float(dia_match.group(1)) if dia_match else None,
        "material": mat_match.group(1).upper() if mat_match else "Unknown",
        "annotated_il_m": float(il_match.group(1)) if il_match else None,
        "annotated_gl_m": float(gl_match.group(1)) if gl_match else None,
    }


def cad_to_real(x: float, y: float, cfg: UUScanConfig) -> tuple[float, float]:
    real_x = x * cfg.unit_to_meters
    real_y = cfg.datum_elevation_m + ((y - cfg.datum_cad_y) * cfg.unit_to_meters / cfg.vertical_exaggeration)
    return real_x, real_y


def interpolate_ground_y(x: float, ground_segments: list[tuple[tuple[float, float], tuple[float, float]]]) -> float | None:
    candidates: list[float] = []
    for (x1, y1), (x2, y2) in ground_segments:
        if min(x1, x2) <= x <= max(x1, x2):
            if math.isclose(x1, x2):
                candidates.append(max(y1, y2))
            else:
                t = (x - x1) / (x2 - x1)
                candidates.append(y1 + t * (y2 - y1))
    return max(candidates) if candidates else None


def scan_uu_section_bytes(file_bytes: bytes, filename: str, cfg: UUScanConfig) -> dict[str, Any]:
    doc = load_cad_from_bytes(file_bytes, filename)
    msp = doc.modelspace()

    pipes: list[dict[str, Any]] = []
    texts: list[dict[str, Any]] = []
    leaders: list[dict[str, Any]] = []
    ground_segments: list[tuple[tuple[float, float], tuple[float, float]]] = []

    for entity in msp:
        etype = entity.dxftype()
        layer = entity.dxf.get("layer", "0")
        handle = entity.dxf.get("handle", "")

        if etype == "CIRCLE":
            cx, cy = cad_to_real(entity.dxf.center.x, entity.dxf.center.y, cfg)
            radius_m = float(entity.dxf.radius) * cfg.unit_to_meters
            pipes.append(
                {
                    "handle": handle,
                    "layer": layer,
                    "center_x_m": round(cx, 4),
                    "center_y_m": round(cy, 4),
                    "drawn_diameter_mm": round(radius_m * 2000.0, 1),
                    "radius_m": radius_m,
                    "matched_text": "",
                    "match_method": "unmatched",
                }
            )
        elif etype in ("LINE", "LWPOLYLINE") and any(k in layer.upper() for k in cfg.ground_layer_keywords):
            if etype == "LINE":
                p1 = cad_to_real(entity.dxf.start.x, entity.dxf.start.y, cfg)
                p2 = cad_to_real(entity.dxf.end.x, entity.dxf.end.y, cfg)
                ground_segments.append((p1, p2))
            else:
                pts = [cad_to_real(p[0], p[1], cfg) for p in entity.get_points("xy")]
                for i in range(len(pts) - 1):
                    ground_segments.append((pts[i], pts[i + 1]))
        elif etype in ("TEXT", "MTEXT"):
            raw = entity.dxf.text if etype == "TEXT" else entity.plain_text()
            clean_text = " ".join(raw.split())
            if not clean_text:
                continue
            ins = entity.dxf.get("insert", (0.0, 0.0, 0.0))
            tx, ty = cad_to_real(ins[0], ins[1], cfg)
            texts.append({"handle": handle, "layer": layer, "text": clean_text, "x_m": tx, "y_m": ty, "assigned": False})
        elif etype == "LEADER":
         verts = list(entity.vertices)
         if len(verts) >= 2:
             tip = cad_to_real(verts[0][0], verts[0][1], cfg)
             tail = cad_to_real(verts[-1][0], verts[-1][1], cfg)
             leaders.append({"tip": tip, "tail": tail})
    # Link LEADERs -> text & pipe
    for ldr in leaders:
        best_text = min(
            (t for t in texts if math.dist(ldr["tail"], (t["x_m"], t["y_m"])) <= cfg.text_search_radius_m),
            key=lambda t: math.dist(ldr["tail"], (t["x_m"], t["y_m"])),
            default=None,
        )
        if not best_text:
            continue
        best_pipe = min(
            (
                p
                for p in pipes
                if abs(math.dist(ldr["tip"], (p["center_x_m"], p["center_y_m"])) - p["radius_m"]) <= cfg.leader_snap_radius_m
            ),
            key=lambda p: abs(math.dist(ldr["tip"], (p["center_x_m"], p["center_y_m"])) - p["radius_m"]),
            default=None,
        )
        if best_pipe:
            best_pipe["matched_text"] = f"{best_pipe['matched_text']} | {best_text['text']}".strip(" |")
            best_pipe["match_method"] = "leader"
            best_text["assigned"] = True

    # Proximity fallback
    for pipe in pipes:
        if pipe["matched_text"]:
            continue
        nearby = sorted(
            (
                (math.dist((pipe["center_x_m"], pipe["center_y_m"]), (t["x_m"], t["y_m"])), t)
                for t in texts
                if not t["assigned"] and math.dist((pipe["center_x_m"], pipe["center_y_m"]), (t["x_m"], t["y_m"])) <= cfg.text_search_radius_m
            ),
            key=lambda item: item[0],
        )
        if nearby:
            pipe["matched_text"] = nearby[0][1]["text"]
            pipe["match_method"] = "proximity"
            nearby[0][1]["assigned"] = True

    enriched_pipes: list[dict[str, Any]] = []
    for idx, p in enumerate(pipes, start=1):
        specs = parse_callout_specs(p["matched_text"])
        eff_dia = specs["annotated_diameter_mm"] or p["drawn_diameter_mm"]
        r_m = (eff_dia / 1000.0) / 2.0
        crown_elev = p["center_y_m"] + r_m
        invert_elev = specs["annotated_il_m"] if specs["annotated_il_m"] is not None else (p["center_y_m"] - r_m)
        ground_elev = specs["annotated_gl_m"] or interpolate_ground_y(p["center_x_m"], ground_segments)
        cover_depth = round(ground_elev - crown_elev, 3) if ground_elev is not None else None

        enriched_pipes.append(
            {
                "pipe_id": f"UU-{idx:02d}",
                "utility_type": classify_utility(p["layer"], p["matched_text"]),
                "diameter_mm": float(eff_dia),
                "material": specs["material"],
                "offset_x_m": float(p["center_x_m"]),
                "center_elev_m": float(p["center_y_m"]),
                "crown_elev_m": round(crown_elev, 3),
                "invert_elev_m": round(invert_elev, 3),
                "ground_elev_m": round(ground_elev, 3) if ground_elev is not None else None,
                "cover_depth_m": cover_depth,
                "layer": p["layer"],
                "matched_text": p["matched_text"],
                "match_method": p["match_method"],
            }
        )

    clearances: list[dict[str, Any]] = []
    for i in range(len(enriched_pipes)):
        for j in range(i + 1, len(enriched_pipes)):
            p1, p2 = enriched_pipes[i], enriched_pipes[j]
            cc_dist = math.dist((p1["offset_x_m"], p1["center_elev_m"]), (p2["offset_x_m"], p2["center_elev_m"]))
            r1, r2 = (p1["diameter_mm"] / 2000.0), (p2["diameter_mm"] / 2000.0)
            gap = round(cc_dist - (r1 + r2), 3)
            status = "CLASH" if gap < 0 else ("LOW_CLEARANCE" if gap < cfg.min_clearance_m else "OK")
            clearances.append(
                {
                    "pipe_a": p1["pipe_id"],
                    "utility_a": p1["utility_type"],
                    "pipe_b": p2["pipe_id"],
                    "utility_b": p2["utility_type"],
                    "center_dist_m": round(cc_dist, 3),
                    "clear_spacing_m": gap,
                    "status": status,
                }
            )

    return {
        "file_name": filename,
        "pipes": enriched_pipes,
        "clearances": clearances,
        "ground_segments": ground_segments,
    }


# ==============================================================================
# 2. TRENCHING & DRAWPIT BOQ COST ENGINE
# ==============================================================================

SURFACE_RATES: dict[str, tuple[float, float]] = {
    "Carriageway - Asphalt": (185.0, 420.0),           # (Breaking HKD/m², Reinstatement HKD/m²)
    "Carriageway - Reinforced Concrete": (340.0, 680.0),
    "Footpath - Concrete Paving Blocks": (95.0, 260.0),
    "Soft Verge / Unpaved": (25.0, 65.0),
}


def build_boq_dataframe(
    pipes_df: pd.DataFrame,
    trench_center_x: float,
    trench_length_m: float,
    trench_width_m: float,
    trench_depth_m: float,
    ground_ref_elev_m: float,
    surface_type: str,
    num_ducts: int,
    duct_dia_mm: int,
    concrete_surround: bool,
    drawpit_count: int,
    pit_l_m: float,
    pit_w_m: float,
    pit_d_m: float,
    pit_type: str,
    uu_buffer_m: float,
    bulking_factor: float,
    tta_days: int,
) -> tuple[pd.DataFrame, list[str]]:
    """Compute engineering quantities and default unit rates for trenching & drawpits."""
    trench_x_min = trench_center_x - (trench_width_m / 2.0)
    trench_x_max = trench_center_x + (trench_width_m / 2.0)
    trench_y_bot = ground_ref_elev_m - trench_depth_m

    # Identify existing UU pipes inside or within proximity buffer of proposed trench box
    conflicting_pipes: list[str] = []
    if not pipes_df.empty:
        for _, row in pipes_df.iterrows():
            r_m = float(row["diameter_mm"]) / 2000.0
            px, py = float(row["offset_x_m"]), float(row["center_elev_m"])
            within_x = (trench_x_min - uu_buffer_m - r_m) <= px <= (trench_x_max + uu_buffer_m + r_m)
            within_y = (trench_y_bot - uu_buffer_m - r_m) <= py <= (ground_ref_elev_m + r_m)
            if within_x and within_y:
                conflicting_pipes.append(f"{row['pipe_id']} ({row['utility_type']} Ø{int(row['diameter_mm'])}mm)")

    # Excavated volumes
    trench_area_m2 = round(trench_length_m * trench_width_m, 2)
    gross_trench_vol_m3 = trench_length_m * trench_width_m * trench_depth_m

    # Hand excavation allowance: 1.5m longitudinal zone per conflicting utility crossing
    hand_dig_length_m = min(trench_length_m, len(conflicting_pipes) * 1.5)
    hand_exc_vol_m3 = round(hand_dig_length_m * trench_width_m * trench_depth_m, 2)
    mech_exc_vol_m3 = round(max(0.0, gross_trench_vol_m3 - hand_exc_vol_m3), 2)

    # Depth-banded mechanical excavation rate
    if trench_depth_m <= 1.5:
        mech_exc_rate = 165.0
        depth_band_label = "0.0m - 1.5m"
    elif trench_depth_m <= 3.0:
        mech_exc_rate = 245.0
        depth_band_label = "1.5m - 3.0m"
    else:
        mech_exc_rate = 360.0
        depth_band_label = "> 3.0m"

    # Shoring wall area (both sides) if depth > 1.2m
    shoring_area_m2 = round(2.0 * trench_length_m * trench_depth_m, 2) if trench_depth_m > 1.2 else 0.0

    # Duct & concrete surround volumes
    total_duct_length_m = round(trench_length_m * num_ducts, 2)
    single_duct_area_m2 = math.pi * ((duct_dia_mm / 2000.0) ** 2)
    ducts_vol_m3 = total_duct_length_m * single_duct_area_m2
    surround_gross_vol_m3 = (trench_length_m * trench_width_m * 0.45) if concrete_surround else 0.0
    concrete_vol_m3 = round(max(0.0, surround_gross_vol_m3 - ducts_vol_m3), 2)

    # Drawpit quantities (external wall thickness assumed 0.25m)
    pit_ext_vol_m3 = round(drawpit_count * (pit_l_m + 0.5) * (pit_w_m + 0.5) * (pit_d_m + 0.25), 2)
    pit_rate = 18500.0 if "Precast" in pit_type else 26000.0

    # Backfill & spoil disposal math
    net_backfill_m3 = round(max(0.0, gross_trench_vol_m3 - ducts_vol_m3 - concrete_vol_m3), 2)
    displaced_spoil_m3 = round((ducts_vol_m3 + concrete_vol_m3 + pit_ext_vol_m3) * bulking_factor, 2)

    break_rate, reinstate_rate = SURFACE_RATES[surface_type]

    items = [
        {
            "Item No": "1.01",
            "Section": "1. Site Clearance & Breaking",
            "Description": f"Saw-cut & break existing {surface_type} along trench alignment",
            "Unit": "m²",
            "Quantity": float(trench_area_m2),
            "Unit Rate (HKD)": float(break_rate),
        },
        {
            "Item No": "2.01",
            "Section": "2. Trench Excavation",
            "Description": f"Mechanical trench excavation in soil (Depth band: {depth_band_label})",
            "Unit": "m³",
            "Quantity": float(mech_exc_vol_m3),
            "Unit Rate (HKD)": float(mech_exc_rate),
        },
        {
            "Item No": "2.02",
            "Section": "2. Trench Excavation",
            "Description": f"Careful hand excavation around {len(conflicting_pipes)} detected live UU crossing(s)",
            "Unit": "m³",
            "Quantity": float(hand_exc_vol_m3),
            "Unit Rate (HKD)": 680.0,
        },
        {
            "Item No": "2.03",
            "Section": "2. Trench Excavation",
            "Description": "Excavate inspection trial pits to expose & verify existing utilities",
            "Unit": "nr",
            "Quantity": float(len(conflicting_pipes)),
            "Unit Rate (HKD)": 3200.0,
        },
        {
            "Item No": "3.01",
            "Section": "3. Lateral Support & Dewatering",
            "Description": "Install & extract steel trench sheets / strutting (required where depth > 1.2m)",
            "Unit": "m²",
            "Quantity": float(shoring_area_m2),
            "Unit Rate (HKD)": 145.0,
        },
        {
            "Item No": "4.01",
            "Section": "4. Duct Laying & Surround",
            "Description": f"Supply & lay DN{duct_dia_mm}mm HDPE corrugated cable/utility duct ({num_ducts}-way)",
            "Unit": "m",
            "Quantity": float(total_duct_length_m),
            "Unit Rate (HKD)": 85.0,
        },
        {
            "Item No": "4.02",
            "Section": "4. Duct Laying & Surround",
            "Description": "Grade 20/20 concrete haunching & surround to ducts incl. marker tape",
            "Unit": "m³",
            "Quantity": float(concrete_vol_m3),
            "Unit Rate (HKD)": 1350.0,
        },
        {
            "Item No": "5.01",
            "Section": "5. Drawpit Construction",
            "Description": f"Excavate for drawpit chambers ({pit_l_m:.1f}m × {pit_w_m:.1f}m × {pit_d_m:.1f}m)",
            "Unit": "m³",
            "Quantity": float(pit_ext_vol_m3),
            "Unit Rate (HKD)": 260.0,
        },
        {
            "Item No": "5.02",
            "Section": "5. Drawpit Construction",
            "Description": f"Construct {pit_type} drawpit incl. sump, cable bearers & heavy-duty DI cover",
            "Unit": "nr",
            "Quantity": float(drawpit_count),
            "Unit Rate (HKD)": float(pit_rate),
        },
        {
            "Item No": "6.01",
            "Section": "6. Backfill & Reinstatement",
            "Description": "Backfill trench with selected granular fill in 200mm compacted layers",
            "Unit": "m³",
            "Quantity": float(net_backfill_m3),
            "Unit Rate (HKD)": 190.0,
        },
        {
            "Item No": "6.02",
            "Section": "6. Backfill & Reinstatement",
            "Description": f"Load & dispose surplus excavated spoil off-site (bulking factor {bulking_factor:.2f}x)",
            "Unit": "m³",
            "Quantity": float(displaced_spoil_m3),
            "Unit Rate (HKD)": 155.0,
        },
        {
            "Item No": "6.03",
            "Section": "6. Backfill & Reinstatement",
            "Description": f"Permanent surface reinstatement ({surface_type})",
            "Unit": "m²",
            "Quantity": float(trench_area_m2),
            "Unit Rate (HKD)": float(reinstate_rate),
        },
        {
            "Item No": "7.01",
            "Section": "7. Traffic Management (TTA)",
            "Description": "Maintain Temporary Traffic Arrangement, barriers, signage & lighting",
            "Unit": "day",
            "Quantity": float(tta_days),
            "Unit Rate (HKD)": 2800.0,
        },
    ]

    df = pd.DataFrame(items)
    df["Amount (HKD)"] = (df["Quantity"] * df["Unit Rate (HKD)"]).round(2)
    return df, conflicting_pipes


# ==============================================================================
# 3. INTERACTIVE PLOTLY CROSS-SECTION RENDERER
# ==============================================================================

def render_cross_section_plot(
    pipes_df: pd.DataFrame,
    ground_segments: list[tuple[tuple[float, float], tuple[float, float]]],
    trench_center_x: float,
    trench_width_m: float,
    trench_depth_m: float,
    ground_ref_elev_m: float,
    uu_buffer_m: float,
) -> go.Figure:
    fig = go.Figure()

    # 1. Proposed Trench Excavation Envelope & Proximity Buffer
    x0 = trench_center_x - trench_width_m / 2.0
    x1 = trench_center_x + trench_width_m / 2.0
    y0 = ground_ref_elev_m - trench_depth_m
    y1 = ground_ref_elev_m

    # Safety buffer zone
    fig.add_shape(
        type="rect",
        x0=x0 - uu_buffer_m,
        y0=y0 - uu_buffer_m,
        x1=x1 + uu_buffer_m,
        y1=y1,
        line=dict(color="rgba(255, 140, 0, 0.8)", width=1.5, dash="dot"),
        fillcolor="rgba(255, 165, 0, 0.08)",
    )
    # Proposed Trench Box
    fig.add_shape(
        type="rect",
        x0=x0,
        y0=y0,
        x1=x1,
        y1=y1,
        line=dict(color="#e63946", width=2.5, dash="dash"),
        fillcolor="rgba(230, 57, 70, 0.16)",
    )
    fig.add_annotation(
        x=trench_center_x,
        y=y0 - 0.18,
        text=f"Proposed Trench ({trench_width_m:.2f}m W × {trench_depth_m:.2f}m D)",
        showarrow=False,
        font=dict(color="#e63946", size=12),
    )

    # 2. Ground Surface Profile
    if ground_segments:
        for idx, ((gx1, gy1), (gx2, gy2)) in enumerate(ground_segments):
            fig.add_trace(
                go.Scatter(
                    x=[gx1, gx2],
                    y=[gy1, gy2],
                    mode="lines",
                    line=dict(color="#2b9348", width=3),
                    name="Existing Ground (GL)" if idx == 0 else None,
                    showlegend=(idx == 0),
                )
            )
    else:
        fig.add_trace(
            go.Scatter(
                x=[x0 - 3.0, x1 + 3.0],
                y=[ground_ref_elev_m, ground_ref_elev_m],
                mode="lines",
                line=dict(color="#2b9348", width=2, dash="dash"),
                name="Reference Ground Level",
            )
        )

    # 3. Existing UU Pipes (drawn as true-scale circles)
    shown_utilities: set[str] = set()
    if not pipes_df.empty:
        for _, row in pipes_df.iterrows():
            cx = float(row["offset_x_m"])
            cy = float(row["center_elev_m"])
            r = float(row["diameter_mm"]) / 2000.0
            util = str(row["utility_type"])
            color = UTILITY_COLORS.get(util, "#7f7f7f")

            fig.add_shape(
                type="circle",
                xref="x",
                yref="y",
                x0=cx - r,
                y0=cy - r,
                x1=cx + r,
                y1=cy + r,
                line=dict(color=color, width=2.5),
                fillcolor=color,
                opacity=0.65,
            )
            fig.add_trace(
                go.Scatter(
                    x=[cx],
                    y=[cy],
                    mode="markers+text",
                    marker=dict(size=6, color=color),
                    text=[f"{row['pipe_id']}<br>Ø{int(row['diameter_mm'])}"],
                    textposition="top center",
                    name=util if util not in shown_utilities else None,
                    showlegend=(util not in shown_utilities),
                    hovertemplate=(
                        f"<b>{row['pipe_id']} — {util}</b><br>"
                        f"Diameter: {int(row['diameter_mm'])} mm<br>"
                        f"Offset X: {cx:.2f} m<br>"
                        f"Invert Level: {row['invert_elev_m']} mPD<br>"
                        f"Cover Depth: {row['cover_depth_m']} m<br>"
                        f"Callout: {row['matched_text']}<extra></extra>"
                    ),
                )
            )
            shown_utilities.add(util)

    fig.update_layout(
        title="Cross-Section Alignment: Existing UU vs. Proposed Trench Envelope",
        xaxis_title="Horizontal Offset (m)",
        yaxis_title="Elevation (mPD)",
        yaxis=dict(scaleanchor="x", scaleratio=1),
        height=460,
        margin=dict(l=40, r=40, t=50, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


# ==============================================================================
# 4. STYLED EXCEL BOQ EXPORTER (WITH LIVE FORMULAS)
# ==============================================================================

def export_boq_to_excel_bytes(
    boq_df: pd.DataFrame,
    pipes_df: pd.DataFrame,
    clearances_df: pd.DataFrame,
    contingency_pct: float,
    project_title: str,
) -> bytes:
    wb = openpyxl.Workbook()
    ws_boq = wb.active
    ws_boq.title = "Priced_BoQ"

    header_fill = PatternFill(start_color="1F3A52", end_color="1F3A52", fill_type="solid")
    total_fill = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    bold_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    ws_boq["A1"] = f"BILL OF QUANTITIES — {project_title.upper()}"
    ws_boq["A1"].font = Font(bold=True, size=14, color="1F3A52")
    ws_boq.merge_cells("A1:G1")

    headers = ["Item No", "Work Section", "Description", "Unit", "Quantity", "Unit Rate (HKD)", "Amount (HKD)"]
    ws_boq.append([])
    ws_boq.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_boq.cell(row=3, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    start_row = 4
    for idx, row in boq_df.iterrows():
        r = start_row + idx
        ws_boq.cell(row=r, column=1, value=row["Item No"]).alignment = Alignment(horizontal="center")
        ws_boq.cell(row=r, column=2, value=row["Section"])
        ws_boq.cell(row=r, column=3, value=row["Description"])
        ws_boq.cell(row=r, column=4, value=row["Unit"]).alignment = Alignment(horizontal="center")

        qty_cell = ws_boq.cell(row=r, column=5, value=float(row["Quantity"]))
        qty_cell.number_format = "#,##0.00"

        rate_cell = ws_boq.cell(row=r, column=6, value=float(row["Unit Rate (HKD)"]))
        rate_cell.number_format = "#,##0.00"

        amt_cell = ws_boq.cell(row=r, column=7, value=f"=ROUND(E{r}*F{r}, 2)")
        amt_cell.number_format = "#,##0.00"

        for c in range(1, 8):
            ws_boq.cell(row=r, column=c).border = thin_border

    end_row = start_row + len(boq_df) - 1
    sub_row = end_row + 1
    cont_row = end_row + 2
    grand_row = end_row + 3

    ws_boq.cell(row=sub_row, column=6, value="Sub-Total (HKD)").font = bold_font
    ws_boq.cell(row=sub_row, column=7, value=f"=SUM(G{start_row}:G{end_row})").font = bold_font
    ws_boq.cell(row=sub_row, column=7).number_format = "#,##0.00"

    ws_boq.cell(row=cont_row, column=6, value=f"Risk Contingency ({contingency_pct:.1f}%)").font = bold_font
    ws_boq.cell(row=cont_row, column=7, value=f"=ROUND(G{sub_row}*{contingency_pct / 100.0}, 2)").font = bold_font
    ws_boq.cell(row=cont_row, column=7).number_format = "#,##0.00"

    ws_boq.cell(row=grand_row, column=6, value="GRAND TOTAL (HKD)").font = Font(bold=True, size=11)
    ws_boq.cell(row=grand_row, column=7, value=f"=G{sub_row}+G{cont_row}").font = Font(bold=True, size=11)
    ws_boq.cell(row=grand_row, column=7).number_format = "#,##0.00"
    for c in range(6, 8):
        ws_boq.cell(row=grand_row, column=c).fill = total_fill
        ws_boq.cell(row=grand_row, column=c).border = thin_border

    col_widths = [12, 26, 54, 10, 14, 18, 18]
    for i, w in enumerate(col_widths, start=1):
        ws_boq.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    def write_df_sheet(sheet_name: str, df: pd.DataFrame) -> None:
        ws = wb.create_sheet(title=sheet_name)
        if df.empty:
            ws["A1"] = "No records extracted."
            return
        ws.append(list(df.columns))
        for c_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = header_fill
            cell.font = header_font
        for r_data in df.itertuples(index=False):
            ws.append(list(r_data))
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(15, len(str(col)) + 5)

    write_df_sheet("Extracted_UU_Pipes", pipes_df)
    write_df_sheet("UU_Clearance_Check", clearances_df)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ==============================================================================
# 5. STREAMLIT WEB INTERFACE
# ==============================================================================

def main() -> None:
    st.set_page_config(page_title="UU Trench & Drawpit Cost Estimator", layout="wide")
    st.title("UU Section Scanner & Drawpit Trenching Cost Estimator")
    st.caption("Upload an AutoCAD UU cross-section (.dxf / .dwg) to extract existing utilities, detect proximity hand-digging zones, and generate a priced BoQ.")

    # Sidebar: CAD Calibration & Trench/Drawpit Design Parameters
    with st.sidebar:
        st.header("1. Project & CAD Calibration")
        project_title = st.text_input("Project Title", value="11kV Cable Trench & Drawpit Installation")
        unit_choice = st.selectbox("CAD Drawing Units", ["Millimeters (mm)", "Meters (m)"])
        unit_to_meters = 0.001 if "mm" in unit_choice else 1.0
        vert_exag = st.number_input("Vertical Exaggeration", min_value=1.0, max_value=20.0, value=1.0, step=0.5)
        datum_cad_y = st.number_input("Datum CAD Y-Coord", value=0.0, step=100.0)
        datum_elev_m = st.number_input("Datum Elevation (mPD)", value=0.0, step=0.5)

        st.divider()
        st.header("2. Proposed Trench & Drawpits")
        trench_center_x = st.number_input("Trench Center Offset X (m)", value=-0.20, step=0.10)
        trench_length_m = st.number_input("Trench Route Length (m)", min_value=1.0, value=120.0, step=5.0)
        trench_width_m = st.number_input("Trench Width (m)", min_value=0.3, value=0.90, step=0.05)
        trench_depth_m = st.number_input("Trench Excavation Depth (m)", min_value=0.4, value=1.45, step=0.05)
        ground_ref_elev_m = st.number_input("Ground Surface Level (mPD)", value=5.00, step=0.10)
        surface_type = st.selectbox("Existing Surface Type", list(SURFACE_RATES.keys()))

        st.subheader("Duct Bank & Drawpits")
        num_ducts = st.number_input("Number of Ducts (ways)", min_value=1, max_value=24, value=4)
        duct_dia_mm = st.selectbox("Duct Diameter (mm)", [110, 150, 160, 200], index=1)
        concrete_surround = st.checkbox("Concrete Surround to Ducts", value=True)
        drawpit_count = st.number_input("Number of Drawpits (nr)", min_value=0, max_value=50, value=3)
        col_p1, col_p2, col_p3 = st.columns(3)
        pit_l_m = col_p1.number_input("Pit L (m)", min_value=0.5, value=1.5, step=0.1)
        pit_w_m = col_p2.number_input("Pit W (m)", min_value=0.5, value=1.2, step=0.1)
        pit_d_m = col_p3.number_input("Pit D (m)", min_value=0.5, value=1.6, step=0.1)
        pit_type = st.selectbox("Drawpit Construction", ["Precast RC Chamber", "Cast-in-Situ RC Chamber"])

        st.divider()
        st.header("3. Risk & Commercial Rules")
        uu_buffer_m = st.number_input("UU Hand-Digging Buffer (m)", min_value=0.1, value=0.50, step=0.10)
        bulking_factor = st.number_input("Spoil Bulking Factor", min_value=1.0, value=1.25, step=0.05)
        tta_days = st.number_input("TTA Duration (days)", min_value=0, value=14, step=1)
        contingency_pct = st.slider("Risk Contingency (%)", min_value=0.0, max_value=30.0, value=10.0, step=1.0)

    # CAD File Upload or Sample Selector
    col_up1, col_up2 = st.columns([3, 1])
    with col_up1:
        uploaded_file = st.file_uploader("Upload UU Section Drawing (.dxf or .dwg)", type=["dxf", "dwg"])
    with col_up2:
        st.write("")
        st.write("")
        use_sample = st.checkbox("Use Built-in Sample UU DXF", value=(uploaded_file is None))

    if uploaded_file is not None:
        cad_bytes = uploaded_file.getvalue()
        cad_filename = uploaded_file.name
    elif use_sample:
        cad_bytes = generate_sample_uu_dxf_bytes()
        cad_filename = "sample_uu_cross_section.dxf"
        st.download_button(
            label="Download Sample DXF to Inspect in AutoCAD",
            data=cad_bytes,
            file_name=cad_filename,
            mime="application/dxf",
        )
    else:
        st.info("Upload a .dxf / .dwg file or check 'Use Built-in Sample UU DXF' above to begin.")
        return

    cfg = UUScanConfig(
        unit_to_meters=unit_to_meters,
        vertical_exaggeration=vert_exag,
        datum_cad_y=datum_cad_y,
        datum_elevation_m=datum_elev_m,
    )

    try:
        scan_result = scan_uu_section_bytes(cad_bytes, cad_filename, cfg)
    except Exception as exc:
        st.error(f"CAD Scan Error: {exc}")
        return

    pipes_df = pd.DataFrame(scan_result["pipes"])
    clearances_df = pd.DataFrame(scan_result["clearances"])

    tab_scan, tab_boq, tab_export = st.tabs(
        ["1. UU CAD Scan & Cross-Section", "2. Interactive BoQ Cost Estimator", "3. Excel Report Export"]
    )

    # TAB 1: CAD Scan & Visual Audit
    with tab_scan:
        clashes = int((clearances_df["status"] == "CLASH").sum()) if not clearances_df.empty else 0
        min_cover = pipes_df["cover_depth_m"].dropna().min() if not pipes_df.empty and pipes_df["cover_depth_m"].notna().any() else None

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Existing UU Pipes Found", len(pipes_df))
        m2.metric("Existing UU Clashes", clashes)
        m3.metric("Min Soil Cover (m)", f"{min_cover:.2f} m" if min_cover is not None else "N/A")
        m4.metric("Trench Cross-Section Area", f"{trench_width_m * trench_depth_m:.2f} m²")

        st.subheader("Extracted Existing Utilities (Editable before Costing)")
        if not pipes_df.empty:
            edited_pipes_df = st.data_editor(
                pipes_df,
                use_container_width=True,
                num_rows="dynamic",
                key="pipes_editor",
            )
        else:
            st.warning("No CIRCLE pipe entities found in modelspace. Check CAD units or drawing layers.")
            edited_pipes_df = pipes_df

        fig = render_cross_section_plot(
            edited_pipes_df,
            scan_result["ground_segments"],
            trench_center_x,
            trench_width_m,
            trench_depth_m,
            ground_ref_elev_m,
            uu_buffer_m,
        )
        st.plotly_chart(fig, use_container_width=True)

        if not clearances_df.empty:
            with st.expander("View Pairwise Existing UU Clearance Matrix"):
                st.dataframe(clearances_df, use_container_width=True)

    # TAB 2: Interactive BoQ Cost Estimator
    with tab_boq:
        boq_df, conflicting_pipes = build_boq_dataframe(
            pipes_df=edited_pipes_df,
            trench_center_x=trench_center_x,
            trench_length_m=trench_length_m,
            trench_width_m=trench_width_m,
            trench_depth_m=trench_depth_m,
            ground_ref_elev_m=ground_ref_elev_m,
            surface_type=surface_type,
            num_ducts=num_ducts,
            duct_dia_mm=duct_dia_mm,
            concrete_surround=concrete_surround,
            drawpit_count=drawpit_count,
            pit_l_m=pit_l_m,
            pit_w_m=pit_w_m,
            pit_d_m=pit_d_m,
            pit_type=pit_type,
            uu_buffer_m=uu_buffer_m,
            bulking_factor=bulking_factor,
            tta_days=tta_days,
        )

        if conflicting_pipes:
            st.warning(
                f"Proximity Alert: {len(conflicting_pipes)} existing utility pipe(s) fall within {uu_buffer_m:.2f}m "
                f"of the proposed trench envelope ({', '.join(conflicting_pipes)}). "
                "Hand excavation and inspection trial pits have been automatically added to the BoQ."
            )
        else:
            st.success("Clear Alignment: No existing utilities detected inside the trench proximity buffer.")

        st.subheader("Priced Bill of Quantities (Edit Quantities or Unit Rates)")
        edited_boq_df = st.data_editor(
            boq_df.drop(columns=["Amount (HKD)"]),
            use_container_width=True,
            disabled=["Item No", "Section", "Unit"],
            key="boq_editor",
        )
        edited_boq_df["Amount (HKD)"] = (edited_boq_df["Quantity"] * edited_boq_df["Unit Rate (HKD)"]).round(2)

        sub_total = float(edited_boq_df["Amount (HKD)"].sum())
        contingency_amt = round(sub_total * (contingency_pct / 100.0), 2)
        grand_total = round(sub_total + contingency_amt, 2)
        cost_per_meter = round(grand_total / trench_length_m, 2)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Base Works Sub-Total", f"HKD {sub_total:,.2f}")
        c2.metric(f"Contingency ({contingency_pct:.0f}%)", f"HKD {contingency_amt:,.2f}")
        c3.metric("Estimated Tender Total", f"HKD {grand_total:,.2f}")
        c4.metric("All-In Rate per Linear Meter", f"HKD {cost_per_meter:,.2f} / m")

        # Cost Breakdown Bar Chart by Section
        sec_summary = edited_boq_df.groupby("Section", as_index=False)["Amount (HKD)"].sum()
        fig_cost = go.Figure(
            go.Bar(
                x=sec_summary["Amount (HKD)"],
                y=sec_summary["Section"],
                orientation="h",
                marker_color="#1f3a52",
                text=[f"HKD {v:,.0f}" for v in sec_summary["Amount (HKD)"]],
                textposition="auto",
            )
        )
        fig_cost.update_layout(
            title="Cost Breakdown by Civil Work Section",
            xaxis_title="Cost (HKD)",
            yaxis=dict(autorange="reversed"),
            height=340,
            margin=dict(l=20, r=20, t=40, b=30),
        )
        st.plotly_chart(fig_cost, use_container_width=True)

    # TAB 3: Excel Export
    with tab_export:
        st.subheader("Download Tender BoQ & CAD Audit Workbook")
        st.write(
            "The exported `.xlsx` workbook contains live Excel formulas (`=ROUND(E*F, 2)` and `=SUM(...)`) "
            "across three tabs: **Priced_BoQ**, **Extracted_UU_Pipes**, and **UU_Clearance_Check**."
        )

        excel_bytes = export_boq_to_excel_bytes(
            boq_df=edited_boq_df,
            pipes_df=edited_pipes_df,
            clearances_df=clearances_df,
            contingency_pct=contingency_pct,
            project_title=project_title,
        )

        st.download_button(
            label="Download Priced BoQ Excel Workbook (.xlsx)",
            data=excel_bytes,
            file_name=f"{Path(cad_filename).stem}_trench_boq_estimate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    main()
