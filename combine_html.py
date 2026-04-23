#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Render exported `layers.json` as a standalone HTML document.

- Text layers become absolutely-positioned HTML blocks with styled spans.
- Image layers reference the exported assets directly.
- Line/geometry layers render with inline SVG.

The output is intended for visual inspection and lightweight preview. It is not a
pixel-perfect PowerPoint renderer, but it preserves slide size, layer order, and
most common styling.
"""

import argparse
import html
import json
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from PIL import ImageFont
except Exception:
    ImageFont = None


_FONT_FILE_REGISTRY: Dict[Tuple[str, str, str, str], Path] = {}

def _group_by_slide(layers: List[Dict[str, Any]]) -> Dict[int, List[Dict[str, Any]]]:
    by_slide: Dict[int, List[Dict[str, Any]]] = defaultdict(list)
    for layer in layers:
        by_slide[int(layer.get("slide", 1))].append(layer)
    for slide_no in by_slide:
        by_slide[slide_no].sort(key=lambda x: int(x.get("layer_index_in_slide", x.get("layer_index_global", 0))))
    return by_slide


def _decode_scaled_protocol_value(value: Any, key: Optional[str] = None, parent_key: Optional[str] = None) -> Any:
    scale = 1000.0
    if isinstance(value, dict):
        current_parent = key if key is not None else parent_key
        decoded = {k: _decode_scaled_protocol_value(v, k, current_parent) for k, v in value.items()}
        keys = set(decoded.keys())
        if keys and keys.issubset({"left", "top", "width", "height"}):
            for k in list(decoded.keys()):
                if isinstance(decoded[k], int):
                    decoded[k] = float(decoded[k]) / scale
        if keys and keys.issubset({"x", "y"}):
            for k in list(decoded.keys()):
                if isinstance(decoded[k], int):
                    decoded[k] = float(decoded[k]) / scale
        return decoded
    if isinstance(value, list):
        decoded_list = [_decode_scaled_protocol_value(v, None, key) for v in value]
        if key and any(str(key).endswith(suffix) for suffix in ("_pt", "_px", "_rel", "_deg")):
            return [float(v) / scale if isinstance(v, int) else v for v in decoded_list]
        return decoded_list
    if isinstance(value, int):
        token = str(key or "")
        parent = str(parent_key or "")
        if token.endswith(("_pt", "_px", "_rel", "_deg")):
            return float(value) / scale
        if token == "rel" and parent in {"line_spacing", "space_before", "space_after"}:
            return float(value) / scale
        if token in {"w", "h"} and parent == "pathLst":
            return float(value) / scale
        if token in {"x", "y"} and parent in {"pt", "pts"}:
            return float(value) / scale
        if token in {"left", "top", "width", "height"} and parent == "box":
            return float(value) / scale
        if token in {"x", "y"} and parent == "pts":
            return float(value) / scale
    return value


def _decode_scaled_protocol_payload(payload: Any) -> Any:
    return _decode_scaled_protocol_value(payload)


def _normalize_layers_payload(payload: Any) -> Tuple[Any, List[Dict[str, Any]]]:
    """
    Support two historical `layers.json` formats:
    1) Flat list of layer dicts, each with optional `slide` and indices.
    2) List-of-slides, where each slide is a list of layer dicts (no `slide` field).

    Returns (raw_payload, flat_layers_for_rendering).
    """
    raw = payload
    if not isinstance(raw, list):
        raise SystemExit("layers.json must contain a JSON array.")

    # Format (2): [ [layer, layer, ...], [layer, ...], ... ]
    if raw and all(isinstance(item, list) for item in raw):
        flat: List[Dict[str, Any]] = []
        global_index = 0
        for slide_idx, slide_layers in enumerate(raw, start=1):
            if not isinstance(slide_layers, list):
                continue
            for layer_idx, layer in enumerate(slide_layers):
                if not isinstance(layer, dict):
                    continue
                normalized = dict(layer)
                normalized.setdefault("slide", slide_idx)
                normalized.setdefault("layer_index_in_slide", layer_idx)
                normalized.setdefault("layer_index_global", global_index)
                global_index += 1
                flat.append(normalized)
        return raw, flat

    # Format (1): [ {layer}, {layer}, ... ]
    flat_layers: List[Dict[str, Any]] = []
    for layer in raw:
        if isinstance(layer, dict):
            flat_layers.append(layer)
    return raw, flat_layers


def _slide_canvas_size(layers: Iterable[Dict[str, Any]]) -> Tuple[int, int]:
    for layer in layers:
        if layer.get("kind") != "slide_canvas":
            continue
        try:
            w = int(layer.get("canvas_width_emu") or 0)
            h = int(layer.get("canvas_height_emu") or 0)
            if w > 0 and h > 0:
                return w, h
        except Exception:
            pass
    return 9144000, 6858000  # 10 x 7.5 inches in EMU


def _box_to_style(box: Optional[Dict[str, Any]]) -> str:
    if not isinstance(box, dict):
        return "left:0;top:0;width:100%;height:100%;"
    left = float(box.get("left", 0.0) or 0.0) * 100.0
    top = float(box.get("top", 0.0) or 0.0) * 100.0
    width = float(box.get("width", 1.0) or 0.0) * 100.0
    height = float(box.get("height", 1.0) or 0.0) * 100.0
    return f"left:{left:.5f}%;top:{top:.5f}%;width:{width:.5f}%;height:{height:.5f}%;"


def _layer_zindex_style(layer: Dict[str, Any]) -> str:
    try:
        idx = int(layer.get("layer_index_in_slide", 0))
        return f"z-index:{100000 - idx}"
    except Exception:
        return "z-index:0"


def _css_color_from_spec(spec: Optional[Dict[str, Any]], default: str = "transparent") -> str:
    if not isinstance(spec, dict):
        return default
    kind = str(spec.get("type") or spec.get("color_type") or "").lower()
    alpha = 1.0
    for mod in spec.get("mods") or []:
        if not isinstance(mod, dict):
            continue
        if str(mod.get("op") or "").lower() != "alpha":
            continue
        try:
            alpha = max(0.0, min(1.0, float(mod.get("val")) / 100000.0))
        except Exception:
            alpha = 1.0
        break

    def _to_css(rgb_value: Any) -> str:
        rgb = str(rgb_value or "").lstrip("#")
        if len(rgb) != 6:
            return default
        if alpha >= 0.999:
            return f"#{rgb}"
        try:
            r = int(rgb[0:2], 16)
            g = int(rgb[2:4], 16)
            b = int(rgb[4:6], 16)
        except Exception:
            return default
        return f"rgba({r}, {g}, {b}, {alpha:.4f})"

    if kind in {"rgb", "srgb"} and spec.get("rgb"):
        return _to_css(spec.get("rgb"))
    if kind == "solid" and str(spec.get("color_type") or "").lower() in {"rgb", "srgb"} and spec.get("rgb"):
        return _to_css(spec.get("rgb"))
    if spec.get("rgb"):
        return _to_css(spec.get("rgb"))
    return default


def _font_weight(value: Any) -> str:
    if value is True:
        return "700"
    if value is False:
        return "400"
    return "inherit"


def _font_style(value: Any) -> str:
    if value is True:
        return "italic"
    if value is False:
        return "normal"
    return "inherit"


def _text_decoration(value: Any) -> str:
    if value is True:
        return "underline"
    if isinstance(value, str) and value not in {"none", "false", ""}:
        return "underline"
    return "none"

def _paragraph_line_scale(para: Dict[str, Any]) -> float:
    _ = para
    return 1.0


def _layer_anchor(layer: Dict[str, Any]) -> str:
    body_pr = layer.get("body_pr")
    if not isinstance(body_pr, dict):
        return "t"
    attrs = body_pr.get("attrs")
    if not isinstance(attrs, dict):
        return "t"
    return str(attrs.get("anchor") or "t").lower()


def _letter_spacing(run: Dict[str, Any]) -> Optional[str]:
    raw = run.get("char_spacing_raw")
    if raw is None:
        return None
    try:
        return f"{float(raw) / 100.0:.2f}pt"
    except Exception:
        return None
def _alignment_css(value: Any) -> str:
    token = str(value or "").split(" ", 1)[0].upper()
    mapping = {
        "LEFT": "left",
        "CENTER": "center",
        "RIGHT": "right",
        "JUSTIFY": "justify",
        "JUSTIFY_LOW": "justify",
        "DISTRIBUTE": "justify",
        "THAI_DISTRIBUTE": "justify",
    }
    return mapping.get(token, "left")


def _vertical_align_css(value: Any) -> str:
    token = str(value or "").split(" ", 1)[0].lower()
    mapping = {
        "t": "flex-start",
        "top": "flex-start",
        "ctr": "center",
        "mid": "center",
        "center": "center",
        "b": "flex-end",
        "bot": "flex-end",
        "bottom": "flex-end",
    }
    return mapping.get(token, "flex-start")


def _vertical_anchor_token(value: Any) -> str:
    token = str(value or "").split(" ", 1)[0].lower()
    if token in {"msoanchormiddle", "middle", "mid", "ctr", "center"}:
        return "ctr"
    if token in {"msoanchorbottom", "bottom", "b", "bot"}:
        return "b"
    return "t"


def _transform_css(layer: Dict[str, Any]) -> str:
    parts: List[str] = []
    if layer.get("flip_h"):
        parts.append("scaleX(-1)")
    if layer.get("flip_v"):
        parts.append("scaleY(-1)")
    rot = float(layer.get("rotation_deg", 0.0) or 0.0)
    if rot:
        parts.append(f"rotate({rot:.4f}deg)")
    if not parts:
        return ""
    return "transform:" + " ".join(parts) + ";transform-origin:center center;"


def _dash_array(dash: Any) -> Optional[str]:
    token = str(dash or "").lower()
    mapping = {
        "solid": None,
        "dash": "12 8",
        "dashdot": "12 6 2 6",
        "lgdash": "18 10",
        "lgdashdot": "18 8 2 8",
        "lgdashdotdot": "18 8 2 8 2 8",
        "sysdash": "10 6",
        "sysdot": "2 6",
        "sysdashdot": "10 6 2 6",
        "sysdashdotdot": "10 6 2 6 2 6",
        "dot": "2 6",
    }
    return mapping.get(token)


def _asset_href(saved_path: Any, html_out: Path, assets_dir: Path) -> str:
    p = Path(str(saved_path or "")).expanduser()
    if not p.is_absolute():
        p = (assets_dir / p.name).resolve()
    if not p.exists():
        alt = (assets_dir / p.name).resolve()
        if alt.exists():
            p = alt
    try:
        rel = os.path.relpath(str(p), str(html_out.parent))
    except Exception:
        rel = str(p)
    return rel.replace(os.sep, "/")


def _style_string(parts: Iterable[str]) -> str:
    cleaned: List[str] = []
    for part in parts:
        text = str(part or "").strip()
        if not text:
            continue
        cleaned.append(text.strip(";"))
    return ";".join(cleaned)


def _safe_name(text: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", text.strip())
    return clean.strip("._") or "font"


def _css_font_family(name: Any) -> str:
    text = str(name or "").strip()
    if not text:
        return "sans-serif"
    return f"{_css_font_face_name(text)}, sans-serif"


def _css_font_face_name(name: Any) -> str:
    text = str(name or "").strip()
    escaped = text.replace("\\", "\\\\").replace("'", "\\'")
    return f"'{escaped}'"


def _font_format_for_suffix(suffix: str) -> str:
    token = suffix.lower()
    if token == ".otf":
        return "opentype"
    if token in {".ttf", ".ttc"}:
        return "truetype"
    return "opentype"


def _font_style_attrs(style_name: str) -> Tuple[str, str]:
    token = str(style_name or "").strip().lower()
    weight = "400"
    style = "normal"
    if "bold" in token:
        weight = "700"
    if "italic" in token or "oblique" in token:
        style = "italic"
    return weight, style


def _font_style_attrs_with_typeface(style_name: str, typeface: str) -> Tuple[str, str]:
    # Some exports encode "Bold" in the typeface name (e.g. "Aileron Bold") while
    # the extracted font part is tagged as "__regular". Prefer matching what runs
    # request via `bold`/`italic` and avoid relying solely on the style tag.
    token = f"{style_name or ''} {typeface or ''}".strip().lower()
    weight, style = _font_style_attrs(style_name)
    if "bold" in token:
        weight = "700"
    if "italic" in token or "oblique" in token:
        style = "italic"
    return weight, style


def _run_weight_style(run: Dict[str, Any]) -> Tuple[str, str]:
    weight = "700" if run.get("bold") is True else "400"
    style = "italic" if run.get("italic") is True else "normal"
    return weight, style


def _extract_webfont(saved_path: Path, out_dir: Path, font_name: str, style_tag: str = "regular") -> Optional[Path]:
    try:
        data = saved_path.read_bytes()
    except Exception:
        return None

    def _wrapped_font_payload(blob: bytes) -> Tuple[bytes, str]:
        # Many PPT embedded fonts are .fntdata parts that wrap sfnt data in an
        # EOT-like container. In that case, the first 4 bytes are the total
        # file size and the next 4 bytes are the raw font-data size.
        if len(blob) >= 8:
            total_size = int.from_bytes(blob[0:4], "little", signed=False)
            font_data_size = int.from_bytes(blob[4:8], "little", signed=False)
            if total_size == len(blob) and 0 < font_data_size <= len(blob):
                start = len(blob) - font_data_size
                payload = blob[start:]
                if payload.startswith(b"OTTO"):
                    return payload, ".otf"
                if payload.startswith(b"ttcf"):
                    return payload, ".ttc"
                if payload.startswith(b"true") or payload.startswith(b"\x00\x01\x00\x00"):
                    return payload, ".ttf"
        return b"", ""

    signatures = [
        (b"OTTO", ".otf"),
        (b"ttcf", ".ttc"),
        (b"true", ".ttf"),
        (b"\x00\x01\x00\x00", ".ttf"),
    ]
    payload, best_ext = _wrapped_font_payload(data)
    if payload:
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{_safe_name(font_name)}__{_safe_name(style_tag)}{best_ext}"
        if out_path.exists():
            return out_path
        out_path.write_bytes(payload)
        return out_path

    best_idx = None
    best_ext = ".ttf"
    for sig, ext in signatures:
        idx = data.find(sig)
        if idx >= 0 and (best_idx is None or idx < best_idx):
            best_idx = idx
            best_ext = ext
    if best_idx is None:
        return None

    payload = data[best_idx:]
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{_safe_name(font_name)}__{_safe_name(style_tag)}{best_ext}"
    if out_path.exists():
        return out_path
    out_path.write_bytes(payload)
    return out_path


def _load_font_faces(layers_json: Path, html_out: Path) -> str:
    fonts_json = layers_json.parent / "fonts.json"
    webfonts_dir = html_out.parent / "webfonts"
    rules: List[str] = []
    seen: set[Tuple[str, str, str]] = set()

    items: List[Dict[str, Any]] = []
    if fonts_json.exists():
        try:
            loaded = json.loads(fonts_json.read_text(encoding="utf-8"))
            if isinstance(loaded, list):
                for item in loaded:
                    if isinstance(item, dict):
                        items.append(item)
        except Exception:
            pass

    fonts_dir = layers_json.parent / "fonts"
    if fonts_dir.exists():
        for p in sorted(fonts_dir.iterdir()):
            if not p.is_file():
                continue
            if p.suffix.lower() not in {".fntdata", ".odttf", ".ttf", ".otf", ".ttc"}:
                continue
            stem = p.stem
            if "__" in stem:
                family, style_name = stem.split("__", 1)
            else:
                family, style_name = stem, "regular"
            items.append(
                {
                    "saved_path": str(p),
                    "typeface": family,
                    "style": style_name,
                    "source": "fonts_dir_scan",
                }
            )

    for item in items:
        if not isinstance(item, dict):
            continue
        typeface = str(item.get("typeface") or "").strip()
        saved_path = str(item.get("saved_path") or "").strip()
        style_name = str(item.get("style") or "regular").strip() or "regular"
        if not typeface or not saved_path:
            continue
        out_font = _extract_webfont(Path(saved_path).expanduser(), webfonts_dir, typeface, style_name)
        if out_font is None:
            continue
        weight, font_style = _font_style_attrs_with_typeface(style_name, typeface)
        key = (typeface, weight, font_style)
        if key in seen:
            continue
        seen.add(key)
        _FONT_FILE_REGISTRY[(str(html_out.parent.resolve()), typeface, weight, font_style)] = out_font
        fmt = _font_format_for_suffix(out_font.suffix)
        rel = os.path.relpath(str(out_font), str(html_out.parent)).replace(os.sep, "/")
        rules.append(
            "\n".join(
                [
                    "@font-face{",
                    f"  font-family:{_css_font_face_name(typeface)};",
                    f"  src:url('{html.escape(rel)}') format('{fmt}');",
                    f"  font-weight:{weight};",
                    f"  font-style:{font_style};",
                    "  font-display:block;",
                    "  font-synthesis:none;",
                    "}",
                ]
            )
        )
    return "\n".join(rules)


def _emu_to_pt(value: Any) -> Optional[float]:
    try:
        return float(value) / 12700.0
    except Exception:
        return None


def _box_height_pt(layer: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> Optional[float]:
    box = layer.get("box")
    if not isinstance(box, dict):
        return None
    try:
        canvas_h_emu = float(canvas_size_emu[1])
        return float(box.get("height", 0.0) or 0.0) * canvas_h_emu / 12700.0
    except Exception:
        return None


def _box_width_pt(layer: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> Optional[float]:
    box = layer.get("box")
    if not isinstance(box, dict):
        return None
    try:
        canvas_w_emu = float(canvas_size_emu[0])
        return float(box.get("width", 0.0) or 0.0) * canvas_w_emu / 12700.0
    except Exception:
        return None


def _body_pr_padding_pt(layer: Dict[str, Any]) -> Tuple[float, float, float, float]:
    body_pr = layer.get("body_pr")
    if not isinstance(body_pr, dict):
        return 0.0, 0.0, 0.0, 0.0
    attrs = body_pr.get("attrs")
    if not isinstance(attrs, dict):
        return 0.0, 0.0, 0.0, 0.0
    left = _emu_to_pt(attrs.get("lIns")) or 0.0
    top = _emu_to_pt(attrs.get("tIns")) or 0.0
    right = _emu_to_pt(attrs.get("rIns")) or 0.0
    bottom = _emu_to_pt(attrs.get("bIns")) or 0.0
    return left, top, right, bottom


def _table_cell_padding_pt(cell: Dict[str, Any]) -> Tuple[float, float, float, float]:
    for keys in (
        ("margin_left_pt", "margin_top_pt", "margin_right_pt", "margin_bottom_pt"),
    ):
        vals = []
        ok = True
        for key in keys:
            v = cell.get(key)
            if v is None:
                ok = False
                break
            try:
                vals.append(float(v))
            except Exception:
                ok = False
                break
        if ok:
            return vals[0], vals[1], vals[2], vals[3]

    body_pr = cell.get("body_pr")
    if isinstance(body_pr, dict) and isinstance(body_pr.get("attrs"), dict):
        attrs = body_pr["attrs"]
        return (
            _emu_to_pt(attrs.get("lIns")) or 0.0,
            _emu_to_pt(attrs.get("tIns")) or 0.0,
            _emu_to_pt(attrs.get("rIns")) or 0.0,
            _emu_to_pt(attrs.get("bIns")) or 0.0,
        )
    return 0.0, 0.0, 0.0, 0.0


def _first_run_color(paragraphs: Any) -> Optional[str]:
    if not isinstance(paragraphs, list):
        return None
    for para in paragraphs:
        if not isinstance(para, dict):
            continue
        for run in para.get("runs") or []:
            if not isinstance(run, dict):
                continue
            color = _css_color_from_spec(run.get("color"), "")
            if color:
                return color
    return None


def _table_cell_align(paragraphs: Any) -> str:
    if not isinstance(paragraphs, list):
        return "left"
    for para in paragraphs:
        if isinstance(para, dict):
            return _alignment_css(para.get("alignment"))
    return "left"


def _table_cell_vertical_anchor(cell: Dict[str, Any]) -> str:
    body_pr = cell.get("body_pr")
    if isinstance(body_pr, dict) and isinstance(body_pr.get("attrs"), dict):
        attrs = body_pr.get("attrs") or {}
        if attrs.get("anchor") is not None:
            return _vertical_anchor_token(attrs.get("anchor"))
    if cell.get("vertical_anchor") is not None:
        return _vertical_anchor_token(cell.get("vertical_anchor"))
    return "t"


def _text_autofit_scale(layer: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> float:
    body_pr = layer.get("body_pr")
    if not isinstance(body_pr, dict) or body_pr.get("autofit") != "spAutoFit":
        return 1.0
    box_h_pt = _box_height_pt(layer, canvas_size_emu)
    if box_h_pt is None or box_h_pt <= 0:
        return 1.0

    total_line_pt = 0.0
    for para in layer.get("paragraphs") or []:
        line_pt = None
        if isinstance(para.get("line_spacing"), dict) and para["line_spacing"].get("mode") == "points":
            try:
                if para["line_spacing"].get("raw") is not None:
                    line_pt = float(para["line_spacing"]["raw"]) / 100.0
                elif para["line_spacing"].get("rel") is not None:
                    line_pt = float(para["line_spacing"]["rel"]) * _canvas_base_pt(canvas_size_emu)
            except Exception:
                line_pt = None
        if line_pt is None:
            max_font = 0.0
            for run in para.get("runs") or []:
                try:
                    max_font = max(max_font, _run_font_size_pt(run, canvas_size_emu))
                except Exception:
                    pass
            line_pt = max_font * 1.2 if max_font > 0 else 0.0
        total_line_pt += line_pt

    if total_line_pt <= 0:
        return 1.0
    height_scale = box_h_pt / total_line_pt
    scale = height_scale
    if scale >= 0.98:
        return 1.0
    return max(0.35, min(1.0, scale))


def _para_spacing_pt(spec: Any) -> float:
    if not isinstance(spec, dict):
        return 0.0
    if str(spec.get("mode") or "").lower() != "points":
        return 0.0
    try:
        return float(spec.get("raw") or 0.0) / 100.0
    except Exception:
        return 0.0


def _paragraph_left_margin_pt(para: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> float:
    attrs = para.get("ppr_attrs")
    if not isinstance(attrs, dict):
        return 0.0
    if attrs.get("marL_rel") is not None:
        try:
            return float(attrs.get("marL_rel")) * _canvas_base_pt(canvas_size_emu)
        except Exception:
            pass
    return _emu_to_pt(attrs.get("marL")) or 0.0


def _paragraph_indent_pt(para: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> float:
    attrs = para.get("ppr_attrs")
    if not isinstance(attrs, dict):
        return 0.0
    if attrs.get("indent_rel") is not None:
        try:
            return float(attrs.get("indent_rel")) * _canvas_base_pt(canvas_size_emu)
        except Exception:
            pass
    return _emu_to_pt(attrs.get("indent")) or 0.0


def _paragraph_bullet_char(para: Dict[str, Any]) -> Optional[str]:
    bullet = para.get("bullet")
    if not isinstance(bullet, dict):
        return None
    if str(bullet.get("type") or "").lower() != "char":
        return None
    char = bullet.get("char")
    return str(char) if char else None


def _paragraph_bullet_font(para: Dict[str, Any]) -> Optional[str]:
    bullet = para.get("bullet")
    if not isinstance(bullet, dict):
        return None
    font = bullet.get("font")
    if not isinstance(font, dict):
        return None
    typeface = font.get("typeface")
    return str(typeface) if typeface else None


def _svg_text_align(para: Dict[str, Any], left: float, right: float) -> Tuple[str, float]:
    align = _alignment_css(para.get("alignment"))
    if align == "center":
        return "middle", (left + right) / 2.0
    if align == "right":
        return "end", right
    return "start", left


def _svg_run_style(run: Dict[str, Any], font_size_pt: float, letter_spacing_pt: Optional[float]) -> str:
    styles = []
    if run.get("font_name"):
        styles.append(f"font-family:{_css_font_family(run['font_name'])}")
    styles.append(f"font-size:{font_size_pt:.2f}px")
    styles.append(f"font-weight:{_font_weight(run.get('bold'))}")
    styles.append(f"font-style:{_font_style(run.get('italic'))}")
    styles.append(f"fill:{_css_color_from_spec(run.get('color'), '#000000')}")
    styles.append("font-kerning:none")
    styles.append("font-variant-ligatures:none")
    styles.append("font-feature-settings:'kern' 0, 'liga' 0, 'clig' 0")
    if letter_spacing_pt is not None:
        styles.append(f"letter-spacing:{letter_spacing_pt:.2f}px")
    if _text_decoration(run.get("underline")) != "none":
        styles.append("text-decoration:underline")
    return ";".join(styles)


def _canvas_base_pt(canvas_size_emu: Tuple[int, int]) -> float:
    try:
        # Use geometric mean baseline: sqrt(w_pt * h_pt). This matches how the
        # rel-only protocol tends to be authored across portrait/square slides.
        w_pt = float(canvas_size_emu[0]) / 12700.0
        h_pt = float(canvas_size_emu[1]) / 12700.0
        return (w_pt * h_pt) ** 0.5
    except Exception:
        return 0.0


def _run_font_size_pt(run: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> float:
    # Newer exports may provide absolute size (`font_size_pt`).
    try:
        val = float(run.get("font_size_pt") or 0.0)
        if val > 0:
            return val
    except Exception:
        pass

    # Older exports may provide relative size (`font_size_rel`) as a fraction of the
    # slide base dimension (min(width_pt, height_pt)).
    rel = run.get("font_size_rel")
    if rel is None:
        return 0.0
    try:
        return float(rel) * _canvas_base_pt(canvas_size_emu)
    except Exception:
        return 0.0


def _run_char_spacing_pt(run: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> Optional[float]:
    # DrawingML char spacing is in 1/100 pt when raw is provided.
    raw = run.get("char_spacing_raw")
    if raw is not None:
        try:
            return float(raw) / 100.0
        except Exception:
            return None
    rel = run.get("char_spacing_rel")
    if rel is None:
        return None
    try:
        return float(rel) * _canvas_base_pt(canvas_size_emu)
    except Exception:
        return None



def _run_font_metrics(
    run: Dict[str, Any],
    autofit_scale: float,
    doc_dir: Path,
    canvas_size_emu: Tuple[int, int],
) -> Tuple[float, Optional[float], Optional[Path]]:
    font_size_pt = _run_font_size_pt(run, canvas_size_emu) * autofit_scale
    letter_spacing_pt: Optional[float] = None
    base_spacing_pt = _run_char_spacing_pt(run, canvas_size_emu)
    if base_spacing_pt is not None:
        letter_spacing_pt = base_spacing_pt * autofit_scale
    font_path = None
    family = str(run.get("font_name") or "").strip()
    if family:
        weight, style = _run_weight_style(run)
        font_path = _FONT_FILE_REGISTRY.get((str(doc_dir.resolve()), family, weight, style))
        if font_path is None:
            font_path = _FONT_FILE_REGISTRY.get((str(doc_dir.resolve()), family, "400", "normal"))
    return font_size_pt, letter_spacing_pt, font_path


def _measure_text_width(text: str, font_size_pt: float, letter_spacing_pt: Optional[float], font_path: Optional[Path]) -> float:
    if not text:
        return 0.0
    width = None
    if ImageFont is not None and font_path is not None:
        try:
            font = ImageFont.truetype(str(font_path), max(1, int(round(font_size_pt))))
            width = float(font.getlength(text))
        except Exception:
            width = None
    if width is None:
        width = font_size_pt * 0.56 * len(text)
    if letter_spacing_pt is not None and len(text) > 1:
        width += (len(text) - 1) * letter_spacing_pt
    return max(0.0, width)


def _measure_font_metrics(font_size_pt: float, font_path: Optional[Path]) -> Tuple[float, float]:
    # NOTE: We intentionally avoid using PIL's font metrics here.
    # For very large display fonts, PIL can return ascent/descent that diverge
    # from how PowerPoint lays out text, which makes baseline-y drift downward.
    _ = font_path
    ascent = font_size_pt * 0.80
    descent = font_size_pt * 0.20
    return ascent, descent


def _tokenize_run_text(text: str) -> List[str]:
    if not text:
        return []
    tokens = re.findall(r"\n|[^\S\n]+|[^\s]+", text)
    return tokens if tokens else [text]


def _wrap_paragraph_runs(
    para: Dict[str, Any],
    autofit_scale: float,
    max_width_pt: float,
    doc_dir: Path,
    canvas_size_emu: Tuple[int, int],
) -> List[List[Dict[str, Any]]]:
    if max_width_pt <= 0:
        return [[]]

    lines: List[List[Dict[str, Any]]] = []
    current_line: List[Dict[str, Any]] = []
    current_width = 0.0

    def flush_line() -> None:
        nonlocal current_line, current_width
        lines.append(current_line)
        current_line = []
        current_width = 0.0

    for run in para.get("runs") or []:
        raw_text = str(run.get("text") or "")
        if not raw_text:
            continue
        font_size_pt, letter_spacing_pt, font_path = _run_font_metrics(run, autofit_scale, doc_dir, canvas_size_emu)
        for token in _tokenize_run_text(raw_text):
            if token == "\n":
                flush_line()
                continue

            token_width = _measure_text_width(token, font_size_pt, letter_spacing_pt, font_path)
            is_space = token.isspace()

            if not current_line:
                if is_space:
                    continue
                current_line.append(
                    {
                        "run": run,
                        "text": token,
                        "font_size_pt": font_size_pt,
                        "letter_spacing_pt": letter_spacing_pt,
                        "font_path": font_path,
                    }
                )
                current_width = token_width
                continue

            if (not is_space) and current_width + token_width > max_width_pt:
                flush_line()
                current_line.append(
                    {
                        "run": run,
                        "text": token,
                        "font_size_pt": font_size_pt,
                        "letter_spacing_pt": letter_spacing_pt,
                        "font_path": font_path,
                    }
                )
                current_width = token_width
                continue

            current_line.append(
                {
                    "run": run,
                    "text": token,
                    "font_size_pt": font_size_pt,
                    "letter_spacing_pt": letter_spacing_pt,
                    "font_path": font_path,
                }
            )
            current_width += token_width

    if current_line or not lines:
        lines.append(current_line)

    # Trim leading/trailing pure-space segments per wrapped line.
    normalized: List[List[Dict[str, Any]]] = []
    for line in lines:
        start = 0
        end = len(line)
        while start < end and str(line[start]["text"]).isspace():
            start += 1
        while end > start and str(line[end - 1]["text"]).isspace():
            end -= 1
        normalized.append(line[start:end])
    return normalized or [[]]


def _merge_line_runs(line_runs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    merged: List[Dict[str, Any]] = []
    for item in line_runs:
        if not merged:
            merged.append(dict(item))
            continue
        prev = merged[-1]
        same_style = (
            prev.get("run") is item.get("run")
            and float(prev.get("font_size_pt") or 0.0) == float(item.get("font_size_pt") or 0.0)
            and (prev.get("letter_spacing_pt") == item.get("letter_spacing_pt"))
            and prev.get("font_path") == item.get("font_path")
        )
        if same_style:
            prev["text"] = f'{prev.get("text") or ""}{item.get("text") or ""}'
        else:
            merged.append(dict(item))
    return merged


def _render_text_layer(layer: Dict[str, Any], canvas_size_emu: Tuple[int, int], doc_dir: Path) -> str:
    box_w_pt = _box_width_pt(layer, canvas_size_emu)
    box_h_pt = _box_height_pt(layer, canvas_size_emu)
    if box_w_pt is None or box_h_pt is None or box_w_pt <= 0 or box_h_pt <= 0:
        return ""

    autofit_scale = _text_autofit_scale(layer, canvas_size_emu)
    container_bits = [
        "position:absolute",
        _box_to_style(layer.get("box")),
        _transform_css(layer),
        _layer_zindex_style(layer),
        "box-sizing:border-box",
        "overflow:visible",
    ]
    fill = _css_color_from_spec(layer.get("shape_fill"))
    line = layer.get("shape_line")
    border_color = "transparent"
    border_w = 0.0
    if isinstance(line, dict):
        border_color = _css_color_from_spec(line.get("color"))
        border_w = float(line.get("width_pt") or 0.0)

    pad_l, pad_t, pad_r, pad_b = _body_pr_padding_pt(layer)
    content_left = pad_l
    content_top = pad_t
    content_right = max(content_left, box_w_pt - pad_r)
    content_bottom = max(content_top, box_h_pt - pad_b)
    available_height = max(0.0, content_bottom - content_top)

    paragraphs_svg: List[str] = []
    para_metrics: List[Tuple[Dict[str, Any], float, float, float, List[List[Dict[str, Any]]]]] = []
    total_content_height = 0.0
    for para in layer.get("paragraphs") or []:
        before_pt = _para_spacing_pt(para.get("space_before"))
        after_pt = _para_spacing_pt(para.get("space_after"))
        para_line_scale = _paragraph_line_scale(para)
        para_left = content_left + _paragraph_left_margin_pt(para, canvas_size_emu)
        para_width = max(1.0, content_right - para_left)
        line_height_pt = 0.0
        if isinstance(para.get("line_spacing"), dict) and para["line_spacing"].get("mode") == "points":
            try:
                if para["line_spacing"].get("raw") is not None:
                    line_height_pt = (float(para["line_spacing"]["raw"]) / 100.0) * autofit_scale * para_line_scale
                elif para["line_spacing"].get("rel") is not None:
                    line_height_pt = float(para["line_spacing"]["rel"]) * _canvas_base_pt(canvas_size_emu) * autofit_scale * para_line_scale
            except Exception:
                line_height_pt = 0.0
        if line_height_pt <= 0:
            max_font_pt = 0.0
            for run in para.get("runs") or []:
                try:
                    max_font_pt = max(max_font_pt, _run_font_size_pt(run, canvas_size_emu) * autofit_scale)
                except Exception:
                    pass
            line_height_pt = max_font_pt * para_line_scale if max_font_pt > 0 else 0.0
        wrapped_lines = _wrap_paragraph_runs(para, autofit_scale, para_width, doc_dir, canvas_size_emu)
        line_count = max(1, len(wrapped_lines))
        para_metrics.append((para, before_pt, line_height_pt, after_pt, wrapped_lines))
        total_content_height += before_pt + line_height_pt * line_count + after_pt

    anchor = _layer_anchor(layer)
    cursor_y = content_top
    if anchor in {"ctr", "mid"} and available_height > total_content_height:
        cursor_y += (available_height - total_content_height) / 2.0
    elif anchor == "b" and available_height > total_content_height:
        cursor_y += available_height - total_content_height

    for para, before_pt, line_height_pt, after_pt, wrapped_lines in para_metrics:
        cursor_y += before_pt
        para_left = content_left + _paragraph_left_margin_pt(para, canvas_size_emu)
        para_right = content_right
        bullet_char = _paragraph_bullet_char(para)
        bullet_font = _paragraph_bullet_font(para)
        bullet_x = para_left + _paragraph_indent_pt(para, canvas_size_emu)
        text_anchor, text_x = _svg_text_align(para, para_left, para_right)
        for line_index, line_runs in enumerate(wrapped_lines):
            line_runs = _merge_line_runs(line_runs)
            max_ascent = 0.0
            max_descent = 0.0
            for item in line_runs:
                try:
                    ascent, descent = _measure_font_metrics(float(item["font_size_pt"]), item.get("font_path"))
                    max_ascent = max(max_ascent, ascent)
                    max_descent = max(max_descent, descent)
                except Exception:
                    pass
            text_block_height = max_ascent + max_descent
            vertical_inset = max(0.0, (line_height_pt - text_block_height) / 2.0)
            # PowerPoint's text layout for very large display fonts tends to behave
            # closer to a line-box top alignment than a pure font-metrics baseline.
            # If ascent exceeds the line height, clamp it to avoid baseline drift.
            effective_ascent = max_ascent
            if line_height_pt > 0:
                effective_ascent = min(effective_ascent, line_height_pt * 0.80)
            baseline_y = cursor_y + vertical_inset + effective_ascent
            runs_svg: List[str] = []
            if line_index == 0 and bullet_char and text_anchor == "start":
                bullet_style = ""
                if para.get("runs"):
                    bullet_run = dict(para["runs"][0])
                    if bullet_font:
                        bullet_run["font_name"] = bullet_font
                    bullet_font_pt, bullet_spacing_pt, _ = _run_font_metrics(bullet_run, autofit_scale, doc_dir, canvas_size_emu)
                    bullet_style = _svg_run_style(bullet_run, bullet_font_pt, bullet_spacing_pt)
                paragraphs_svg.append(
                    f'<text x="{bullet_x:.4f}" y="{baseline_y:.4f}" text-anchor="start" '
                    f'dominant-baseline="alphabetic" xml:space="preserve" style="{bullet_style}">{html.escape(bullet_char)}</text>'
                )
            for item in line_runs:
                run = item["run"]
                text = html.escape(str(item["text"] or ""))
                if not text:
                    continue
                runs_svg.append(
                    f'<tspan style="{_svg_run_style(run, float(item["font_size_pt"]), item["letter_spacing_pt"])}">{text}</tspan>'
                )
            paragraphs_svg.append(
                f'<text x="{text_x:.4f}" y="{baseline_y:.4f}" text-anchor="{text_anchor}" '
                f'dominant-baseline="alphabetic" xml:space="preserve">{"".join(runs_svg)}</text>'
            )
            cursor_y += line_height_pt
        cursor_y += after_pt

    bg_rect = ""
    if fill != "transparent" or (border_color != "transparent" and border_w > 0):
        inset = max(0.0, border_w / 2.0)
        width = max(0.0, box_w_pt - inset * 2.0)
        height = max(0.0, box_h_pt - inset * 2.0)
        bg_rect = (
            f'<rect x="{inset:.4f}" y="{inset:.4f}" width="{width:.4f}" height="{height:.4f}" '
            f'fill="{"none" if fill == "transparent" else fill}" '
            f'stroke="{"none" if border_color == "transparent" else border_color}" '
            f'stroke-width="{border_w:.2f}" vector-effect="non-scaling-stroke"/>'
        )

    parts = [
        (
            f'<svg class="layer text-layer" data-kind="text" style="{_style_string(container_bits)}" '
            f'viewBox="0 0 {box_w_pt:.4f} {box_h_pt:.4f}" preserveAspectRatio="none">'
        )
    ]
    if bg_rect:
        parts.append(f"  {bg_rect}")
    parts.extend(f"  {node}" for node in paragraphs_svg)
    parts.append("</svg>")
    return "\n".join(parts)


def _render_image_layer(layer: Dict[str, Any], html_out: Path, assets_dir: Path) -> str:
    src = _asset_href(layer.get("saved_path"), html_out, assets_dir)
    style = _style_string(
        [
            "position:absolute",
            _box_to_style(layer.get("box")),
            _transform_css(layer),
            _layer_zindex_style(layer),
            "overflow:hidden",
        ]
    )
    return "\n".join(
        [
            f'<div class="layer image-layer" data-kind="{html.escape(str(layer.get("kind") or "image"))}" style="{style}">',
            f'  <img src="{html.escape(src)}" alt="" style="width:100%;height:100%;display:block;object-fit:fill;"/>',
            "</div>",
        ]
    )


def _render_table_layer(layer: Dict[str, Any], canvas_size_emu: Tuple[int, int]) -> str:
    style = _style_string(
        [
            "position:absolute",
            _box_to_style(layer.get("box")),
            _transform_css(layer),
            _layer_zindex_style(layer),
            "overflow:hidden",
        ]
    )
    rows = layer.get("cells") or []
    col_widths = layer.get("col_widths_pt") or []
    col_total = sum(float(v or 0.0) for v in col_widths) or 1.0

    colgroup = []
    for w in col_widths:
        try:
            pct = (float(w or 0.0) / col_total) * 100.0
        except Exception:
            pct = 0.0
        colgroup.append(f'    <col style="width:{pct:.5f}%"/>')

    tr_nodes: List[str] = []
    for row in rows:
        if not isinstance(row, list):
            continue
        td_nodes: List[str] = []
        for cell in row:
            if not isinstance(cell, dict):
                continue
            fill = _css_color_from_spec(cell.get("fill"), "transparent")
            borders = cell.get("borders") or {}
            border_parts = []
            for side_key, css_side in (("l", "left"), ("r", "right"), ("t", "top"), ("b", "bottom")):
                spec = borders.get(side_key)
                if not isinstance(spec, dict):
                    continue
                color = _css_color_from_spec(spec.get("color"), "#000000")
                width = float(spec.get("width_pt") or 1.0)
                border_parts.append(f"border-{css_side}:{width:.2f}pt solid {color}")
            pad_l, pad_t, pad_r, pad_b = _table_cell_padding_pt(cell)
            text_align = _table_cell_align(cell.get("paragraphs"))
            vertical_align = _vertical_align_css(_table_cell_vertical_anchor(cell))
            text_color = _first_run_color(cell.get("paragraphs"))
            td_style = _style_string(
                [
                    f"background:{fill}" if fill != "transparent" else "",
                    *border_parts,
                    "padding:0",
                ]
            )
            inner_style = _style_string(
                [
                    "width:100%",
                    "height:100%",
                    "display:flex",
                    f"align-items:{vertical_align}",
                    "box-sizing:border-box",
                    f"padding:{pad_t:.2f}pt {pad_r:.2f}pt {pad_b:.2f}pt {pad_l:.2f}pt",
                    f"text-align:{text_align}",
                    f"color:{text_color}" if text_color else "",
                ]
            )
            text = html.escape(str(cell.get("text") or "")).replace("\n", "<br/>")
            td_nodes.append(f'      <td style="{td_style}"><div style="{inner_style}">{text}</div></td>')
        tr_nodes.append("\n".join(["    <tr>", *td_nodes, "    </tr>"]))

    return "\n".join(
        [
            f'<div class="layer table-layer" data-kind="ppt_graph_table" style="{style}">',
            '  <table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed;">',
            "  <colgroup>",
            *colgroup,
            "  </colgroup>",
            "  <tbody>",
            *tr_nodes,
            "  </tbody>",
            "  </table>",
            "</div>",
        ]
    )


def _marker_defs() -> str:
    return (
        '<defs>'
        '<marker id="arrow-end" markerWidth="10" markerHeight="10" refX="8" refY="5" orient="auto" markerUnits="strokeWidth">'
        '<path d="M0,0 L10,5 L0,10 z" fill="context-stroke"/></marker>'
        '<marker id="arrow-start" markerWidth="10" markerHeight="10" refX="2" refY="5" orient="auto" markerUnits="strokeWidth">'
        '<path d="M10,0 L0,5 L10,10 z" fill="context-stroke"/></marker>'
        "</defs>"
    )


def _line_marker_attr(end_spec: Any, side: str) -> str:
    if not isinstance(end_spec, dict):
        return ""
    kind = str(end_spec.get("type") or "").lower()
    if kind in {"none", ""}:
        return ""
    return f' marker-{side}="url(#{ "arrow-start" if side == "start" else "arrow-end" })"'


def _render_line_layer(layer: Dict[str, Any]) -> str:
    spec = layer.get("shape_xml") or {}
    p1 = spec.get("p1") or {}
    p2 = spec.get("p2") or {}
    line = spec.get("line") or {}
    stroke = _css_color_from_spec(line.get("color"), "#000000")
    width_pt = float(line.get("width_pt") or 1.0)
    dash_array = _dash_array(line.get("dash"))
    dash_attr = f' stroke-dasharray="{dash_array}"' if dash_array else ""
    marker_start = _line_marker_attr(line.get("head_end"), "start")
    marker_end = _line_marker_attr(line.get("tail_end"), "end")
    line_tag = (
        f'<line x1="{float(p1.get("x", 0.0)) * 1000:.4f}" y1="{float(p1.get("y", 0.0)) * 1000:.4f}" '
        f'x2="{float(p2.get("x", 0.0)) * 1000:.4f}" y2="{float(p2.get("y", 0.0)) * 1000:.4f}" '
        f'stroke="{stroke}" stroke-width="{width_pt:.2f}" vector-effect="non-scaling-stroke" '
        f'stroke-linecap="round" stroke-linejoin="round"{dash_attr}{marker_start}{marker_end}/>'
    )
    return "\n".join(
        [
            (
                '<svg class="layer line-layer" data-kind="ppt_graph_line" '
                f'style="{_style_string(["position:absolute", "left:0", "top:0", "width:100%", "height:100%", "overflow:visible", _layer_zindex_style(layer)])}" '
                'viewBox="0 0 1000 1000" preserveAspectRatio="none">'
            ),
            f"  {_marker_defs()}",
            f"  {line_tag}",
            "</svg>",
        ]
    )


def _svg_path_from_commands(commands: List[Dict[str, Any]], path_w: float, path_h: float) -> str:
    parts: List[str] = []
    for cmd in commands:
        ctype = cmd.get("type")
        if ctype in {"moveTo", "lnTo"} and isinstance(cmd.get("pt"), dict):
            pt = cmd["pt"]
            x = float(pt.get("x") or 0.0) * path_w
            y = float(pt.get("y") or 0.0) * path_h
            parts.append(("M" if ctype == "moveTo" else "L") + f" {x:.4f} {y:.4f}")
        elif ctype == "cubicBezTo" and isinstance(cmd.get("pts"), list) and len(cmd["pts"]) == 3:
            nums: List[str] = []
            for pt in cmd["pts"]:
                x = float(pt.get("x") or 0.0) * path_w
                y = float(pt.get("y") or 0.0) * path_h
                nums.append(f"{x:.4f} {y:.4f}")
            parts.append("C " + " ".join(nums))
        elif ctype == "close":
            parts.append("Z")
    return " ".join(parts)


def _is_normalized_custom_path(path: Dict[str, Any]) -> bool:
    pts: List[Dict[str, Any]] = []
    for cmd in path.get("commands") or []:
        if not isinstance(cmd, dict):
            continue
        if isinstance(cmd.get("pt"), dict):
            pts.append(cmd["pt"])
        elif isinstance(cmd.get("pts"), list):
            for pt in cmd["pts"]:
                if isinstance(pt, dict):
                    pts.append(pt)
    if not pts:
        return False
    try:
        max_x = max(float(pt.get("x") or 0.0) for pt in pts)
        max_y = max(float(pt.get("y") or 0.0) for pt in pts)
        path_w = float(path.get("w") or 1.0)
        path_h = float(path.get("h") or 1.0)
    except Exception:
        return False
    return max_x <= 1.001 and max_y <= 1.001 and path_w <= 1.001 and path_h <= 1.001


def _render_geo_layer(layer: Dict[str, Any]) -> str:
    box_style = _box_to_style(layer.get("box"))
    transform = _transform_css(layer)
    fill = _css_color_from_spec(layer.get("fill"))
    line = layer.get("line") or {}
    stroke = _css_color_from_spec(line.get("color"), "transparent")
    stroke_width = float(line.get("width_pt") or 0.0)
    dash_array = _dash_array(line.get("dash"))
    dash_attr = f' stroke-dasharray="{dash_array}"' if dash_array else ""
    shape_xml = layer.get("shape_xml") or {}
    fill_attr = "none" if fill == "transparent" else fill
    stroke_attr = "none" if stroke == "transparent" else stroke
    svg_style = _style_string(["position:absolute", box_style, transform, _layer_zindex_style(layer), "overflow:visible"])

    if isinstance(shape_xml, dict) and shape_xml.get("format") == "shape_spec_v1" and shape_xml.get("type") == "geo":
        geom = shape_xml.get("geom") or {}
        if isinstance(geom, dict) and geom.get("type") == "prstGeom":
            prst = str(geom.get("prst") or "").lower()
            inset = max(0.0, stroke_width / 2.0)
            size = max(0.0, 100.0 - inset * 2.0)
            if prst == "ellipse":
                shape_tag = f'<ellipse cx="50" cy="50" rx="{max(0.0, 50.0 - inset):.4f}" ry="{max(0.0, 50.0 - inset):.4f}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_width:.2f}" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
            elif prst == "roundrect":
                radius = max(0.0, 12.0 - inset)
                shape_tag = f'<rect x="{inset:.4f}" y="{inset:.4f}" width="{size:.4f}" height="{size:.4f}" rx="{radius:.4f}" ry="{radius:.4f}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_width:.2f}" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
            else:
                shape_tag = f'<rect x="{inset:.4f}" y="{inset:.4f}" width="{size:.4f}" height="{size:.4f}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_width:.2f}" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
            return (
                f'<svg class="layer geo-layer" data-kind="ppt_graph_geo" style="{svg_style}" '
                'viewBox="0 0 100 100" preserveAspectRatio="none">'
                + shape_tag
                + "</svg>"
            )
        if isinstance(geom, dict) and geom.get("type") == "custGeom":
            paths = geom.get("paths") or []
            svg_paths: List[str] = []
            view_w = 1.0
            view_h = 1.0
            for path in paths:
                if not isinstance(path, dict):
                    continue
                if _is_normalized_custom_path(path):
                    path_w = 1.0
                    path_h = 1.0
                else:
                    path_w = max(1.0, float(path.get("w") or 0.0))
                    path_h = max(1.0, float(path.get("h") or 0.0))
                view_w = max(view_w, path_w)
                view_h = max(view_h, path_h)
                commands = []
                for cmd in path.get("commands") or []:
                    if not isinstance(cmd, dict):
                        continue
                    op = cmd.get("op")
                    pts = cmd.get("pts") or []
                    if op in {"moveTo", "lnTo"} and pts:
                        pt = pts[0]
                        x = float(pt.get("x") or 0.0) * path_w
                        y = float(pt.get("y") or 0.0) * path_h
                        commands.append(("M" if op == "moveTo" else "L") + f" {x:.4f} {y:.4f}")
                    elif op == "cubicBezTo" and len(pts) == 3:
                        nums: List[str] = []
                        for pt in pts:
                            x = float(pt.get("x") or 0.0) * path_w
                            y = float(pt.get("y") or 0.0) * path_h
                            nums.append(f"{x:.4f} {y:.4f}")
                        commands.append("C " + " ".join(nums))
                    elif op == "close":
                        commands.append("Z")
                d = " ".join(commands)
                if d:
                    svg_paths.append(
                        f'<path d="{d}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_width:.2f}" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
                    )
            if svg_paths:
                return (
                    f'<svg class="layer geo-layer" data-kind="ppt_graph_geo" style="{svg_style}" '
                    f'viewBox="0 0 {view_w:.4f} {view_h:.4f}" preserveAspectRatio="none">'
                    + "".join(svg_paths)
                    + "</svg>"
                )

    if isinstance(shape_xml, dict) and isinstance(shape_xml.get("spPr"), dict):
        cust = (shape_xml.get("spPr") or {}).get("custGeom") or {}
        paths = cust.get("pathLst") or []
        svg_paths: List[str] = []
        view_w = 1.0
        view_h = 1.0
        for path in paths:
            if not isinstance(path, dict):
                continue
            if _is_normalized_custom_path(path):
                path_w = 1.0
                path_h = 1.0
            else:
                path_w = max(1.0, float(path.get("w") or 0.0))
                path_h = max(1.0, float(path.get("h") or 0.0))
            view_w = max(view_w, path_w)
            view_h = max(view_h, path_h)
            d = _svg_path_from_commands(path.get("commands") or [], path_w, path_h)
            if d:
                svg_paths.append(
                        f'<path d="{d}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_width:.2f}" vector-effect="non-scaling-stroke" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
                )
        if svg_paths:
            return (
                f'<svg class="layer geo-layer" data-kind="ppt_graph_geo" style="{svg_style}" '
                f'viewBox="0 0 {view_w:.4f} {view_h:.4f}" preserveAspectRatio="none">'
                + "".join(svg_paths)
                + "</svg>"
            )

    return (
        f'<div class="layer geo-fallback" data-kind="ppt_graph_geo" style="{_style_string(["position:absolute", box_style, transform, _layer_zindex_style(layer), f"background:{fill}", f"border:{stroke_width:.2f}pt solid {stroke}", "box-sizing:border-box"])}"></div>'
    )


def _render_layer(
    layer: Dict[str, Any],
    html_out: Path,
    assets_dir: Path,
    canvas_size_emu: Tuple[int, int],
) -> str:
    kind = layer.get("kind")
    if kind == "text":
        return _render_text_layer(layer, canvas_size_emu, html_out.parent)
    if kind == "ppt_graph_table":
        return _render_table_layer(layer, canvas_size_emu)
    if isinstance(kind, str) and ("image" in kind):
        return _render_image_layer(layer, html_out, assets_dir)
    if kind == "ppt_graph_line":
        return _render_line_layer(layer)
    if kind == "ppt_graph_geo":
        return _render_geo_layer(layer)
    return ""


def _slide_background_css(slide_layers: List[Dict[str, Any]]) -> str:
    for layer in slide_layers:
        if layer.get("kind") != "slide_canvas":
            continue
        bg = layer.get("background_fill")
        color = _css_color_from_spec(bg, "#ffffff")
        return f"background:{color};"
    return "background:#ffffff;"


def _render_slide(
    slide_no: int,
    slide_layers: List[Dict[str, Any]],
    html_out: Path,
    assets_dir: Path,
    canvas_size_emu: Tuple[int, int],
) -> str:
    canvas_w, canvas_h = canvas_size_emu
    body: List[str] = []
    for layer in reversed(slide_layers):
        if layer.get("kind") == "slide_canvas":
            continue
        rendered = _render_layer(layer, html_out, assets_dir, canvas_size_emu)
        if rendered:
            body.append(rendered)
    bg_css = _slide_background_css(slide_layers)
    lines = [
        '<section class="slide-wrap">',
        f'  <div class="slide" style="aspect-ratio:{canvas_w} / {canvas_h};{bg_css}">',
    ]
    lines.extend(f"    {node.replace(chr(10), chr(10) + '    ')}" for node in body)
    lines.extend(["  </div>", "</section>"])
    return "\n".join(lines)


def build_html(layers_json: Path, assets_dir: Path, out_html: Path) -> None:
    raw_payload = _decode_scaled_protocol_payload(json.loads(layers_json.read_text(encoding="utf-8")))
    raw_payload, layers = _normalize_layers_payload(raw_payload)
    by_slide = _group_by_slide(layers)
    canvas_size = _slide_canvas_size(layers)
    font_face_css = _load_font_faces(layers_json, out_html)

    slides_html = [
        _render_slide(slide_no, by_slide[slide_no], out_html, assets_dir, canvas_size)
        for slide_no in sorted(by_slide.keys())
    ]

    title = html.escape(layers_json.parent.name or "slides")
    slides_block = "\n".join("    " + slide.replace("\n", "\n    ") for slide in slides_html)
    # Embed the exact layer protocol in the HTML so it can be reconstructed later
    # (e.g. by test_html.py) without lossy parsing.
    embedded_layers_json = json.dumps(raw_payload, ensure_ascii=False, indent=2).replace("</", "<\\/")
    doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>{title}</title>
  <style>
    {font_face_css}
    :root {{
      color-scheme: light;
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: #f3f4f6;
      color: #111827;
      padding: 24px;
    }}
    .deck {{
      display: flex;
      flex-direction: column;
      gap: 28px;
      align-items: center;
    }}
    .slide-wrap {{
      width: min(96vw, 1280px);
    }}
    .slide {{
      container-type: size;
      position: relative;
      width: 100%;
      overflow: hidden;
      box-shadow: 0 8px 30px rgba(0, 0, 0, 0.12);
      border-radius: 6px;
    }}
    .layer {{
      position: absolute;
    }}
    .text-layer {{
      white-space: normal;
    }}
    .text-layer p {{
      margin: 0;
    }}
    .geo-layer, .line-layer {{
      overflow: visible;
    }}
  </style>
</head>
<body>
  <main class="deck">
{slides_block}
  </main>
  <script id="layers-json" type="application/json">
{embedded_layers_json}
  </script>
</body>
</html>
"""
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(doc, encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(description="Render layers.json to a preview HTML file.")
    ap.add_argument("layers_json", help="Path to layers.json")
    ap.add_argument("--assets", default=None, help="Assets folder (default: <layers_json_dir>/assets)")
    ap.add_argument(
        "-o",
        "--out",
        default=None,
        help="Output HTML path (default: <layers_json_dir>/rebuilt.html)",
    )
    args = ap.parse_args()

    layers_json = Path(args.layers_json).expanduser().resolve()
    assets_dir = (
        Path(args.assets).expanduser().resolve()
        if args.assets
        else (layers_json.parent / "assets").resolve()
    )
    out_html = (
        Path(args.out).expanduser().resolve()
        if args.out
        else (layers_json.parent / "rebuilt.html").resolve()
    )
    build_html(layers_json, assets_dir, out_html)
    print(f"Saved: {out_html}")


if __name__ == "__main__":
    main()
