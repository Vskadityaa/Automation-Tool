from flask import Flask, request, send_file, render_template, session, redirect, url_for
import pandas as pd
import os, uuid, re, base64
import xml.etree.ElementTree as ET

try:
    import cairosvg
    HAS_CAIROSVG = True
except (ImportError, OSError):
    HAS_CAIROSVG = False

# =================================================
# SVG NAMESPACE
# =================================================
SVG_NS = "http://www.w3.org/2000/svg"
XLINK_NS = "http://www.w3.org/1999/xlink"
VISIO_NS = "http://schemas.microsoft.com/visio/2003/SVGExtensions/"
ET.register_namespace("", SVG_NS)


def svg_tag(tag):
    return f"{{{SVG_NS}}}{tag}"

def get_viewbox(svg_root):
    vb = svg_root.attrib.get("viewBox")
    if vb:
        p = vb.replace(",", " ").split()
        if len(p) == 4:
            return float(p[2]), float(p[3])
    return 800, 600

# =================================================
# APP + PATHS
# =================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "bms-merge-dashboard-secret")
# When hosted behind a reverse proxy (Render, Railway, etc.), use real public URL
try:
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
UP = os.path.join(BASE, "uploads")
EXCEL_DIR = os.path.join(UP, "excel")
DRAWING_DIR = os.path.join(UP, "drawing")
TEMP_DIR = os.path.join(UP, "temp")
SVG_TEMPLATES_DIR = os.path.join(UP, "svg_templates")

for d in (EXCEL_DIR, DRAWING_DIR, TEMP_DIR, SVG_TEMPLATES_DIR):
    os.makedirs(d, exist_ok=True)


def list_svg_templates():
    """Return list of (filename, display_name) for saved SVG templates."""
    if not os.path.isdir(SVG_TEMPLATES_DIR):
        return []
    out = []
    for f in sorted(os.listdir(SVG_TEMPLATES_DIR)):
        if f.endswith(".svg"):
            name = os.path.splitext(f)[0].replace("_", " ")
            out.append((f, name))
    return out

# =================================================
# XML SAFE
# =================================================
def xml_escape(val):
    if val is None:
        return ""
    return (
        str(val)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )

# =================================================
# HEADER DETECTION
# =================================================
def is_table_header(row):
    text = " ".join(str(v).lower() for v in row if pd.notna(v))
    required = ["software", "system", "object", "description"]
    return all(k in text for k in required)

# =================================================
# READ TABLES
# =================================================
def read_all_tables(df):
    tables = []
    current_rows = []
    in_table = False
    blank_count = 0

    for _, row in df.iterrows():
        is_blank = all(pd.isna(v) or str(v).strip() == "" for v in row)

        if is_table_header(row):
            if current_rows:
                tables.append(pd.DataFrame(current_rows))
                current_rows = []
            in_table = True
            blank_count = 0
            continue

        if not in_table:
            continue

        if is_blank:
            blank_count += 1
        else:
            blank_count = 0

        if blank_count >= 2:
            tables.append(pd.DataFrame(current_rows))
            current_rows = []
            in_table = False
            blank_count = 0
            continue

        current_rows.append(row)

    if current_rows:
        tables.append(pd.DataFrame(current_rows))

    return tables

# =================================================
# SAFE COLUMN
# =================================================
def safe_col(df, idx):
    if idx < df.shape[1]:
        return df.iloc[:, idx]
    return pd.Series([""] * len(df))

# =================================================
# TEXT WRAP
# =================================================
def wrap_text(text, max_chars):
    words = str(text).split()
    lines, line = [], ""
    for w in words:
        if len(line + w) <= max_chars:
            line += w + " "
        else:
            lines.append(line.strip())
            line = w + " "
    if line:
        lines.append(line.strip())
    return lines

# =================================================
# BUILD SVG TABLE
# =================================================
def build_table_svg(df):
    font_size = 10
    line_h = 14
    padding = 6
    start_y = 30

    col_x = [20, 110, 270, 430, 690]
    col_w = [90, 160, 160, 260, 160]
    wrap_limits = [8, 18, 16, 28, 14]

    rows = []
    for _, r in df.iterrows():
        cells = [
            wrap_text(r["POINT"], wrap_limits[0]),
            wrap_text(r["SYSTEM"], wrap_limits[1]),
            wrap_text(r["OBJECT"], wrap_limits[2]),
            wrap_text(r["DESCRIPTION"], wrap_limits[3]),
            wrap_text(r["SIGNAL"], wrap_limits[4]),
        ]
        max_lines = max(len(c) for c in cells)
        rows.append((cells, max_lines * line_h + padding * 2))

    height = start_y + sum(rh for _, rh in rows) + 40
    width = max(col_x) + max(col_w) + 20

    svg = [
        f'<svg xmlns="{SVG_NS}" width="{width}" height="{height}">',
        '<style>text{font-family:Arial;}</style>'
    ]

    headers = ["POINT", "SYSTEM", "OBJECT", "DESCRIPTION", "SIGNAL"]
    y = start_y

    for i, h in enumerate(headers):
        svg.append(
            f'<rect x="{col_x[i]}" y="{y}" width="{col_w[i]}" height="28" fill="#e5e7eb" stroke="black"/>'
        )
        svg.append(
            f'<text x="{col_x[i]+padding}" y="{y+18}" font-size="{font_size}" font-weight="bold">{h}</text>'
        )

    y += 28

    for cells, rh in rows:
        for i, cell in enumerate(cells):
            svg.append(
                f'<rect x="{col_x[i]}" y="{y}" width="{col_w[i]}" height="{rh}" fill="white" stroke="black"/>'
            )
            ty = y + padding + font_size
            for line in cell:
                svg.append(
                    f'<text x="{col_x[i]+padding}" y="{ty}" font-size="{font_size}">{xml_escape(line)}</text>'
                )
                ty += line_h
        y += rh

    svg.append("</svg>")
    return "\n".join(svg)

# =================================================
# COMPACT EXCEL TABLE (merge into drawing)
# =================================================
def append_full_excel_table(root, df):
    if df.empty:
        return
    if "viewBox" in root.attrib:
        vb = list(map(float, root.attrib["viewBox"].split()))
        svg_width = vb[2]
    else:
        svg_width = float(root.attrib.get("width", 1600))
    row_height = 18
    header_height = 20
    title_height = 22
    padding = 4
    columns = list(df.columns)
    col_widths = []
    for col in columns:
        max_len = max(
            [len(str(col))] + [len(str(v)) for v in df[col].astype(str)]
        )
        width = max(70, min(max_len * 5, 160))
        col_widths.append(width)
    table_width = sum(col_widths)
    table_height = title_height + header_height + (len(df) * row_height)
    start_x = svg_width - table_width - 40
    start_y = 50
    table = ET.SubElement(root, f"{{{SVG_NS}}}g", {
        "transform": f"translate({start_x},{start_y})"
    })
    ET.SubElement(table, f"{{{SVG_NS}}}rect", {
        "x": "0", "y": "0", "width": str(table_width), "height": str(table_height),
        "fill": "#fff", "stroke": "#000", "stroke-width": "1"
    })
    ET.SubElement(table, f"{{{SVG_NS}}}text", {
        "x": str(table_width / 2), "y": "14", "text-anchor": "middle",
        "font-size": "9", "font-family": "Arial", "font-weight": "bold"
    }).text = "Excel Data"
    ET.SubElement(table, f"{{{SVG_NS}}}rect", {
        "x": "0", "y": str(title_height), "width": str(table_width),
        "height": str(header_height), "fill": "#f2f2f2", "stroke": "#000"
    })
    x_cursor = 0
    for i, col in enumerate(columns):
        ET.SubElement(table, f"{{{SVG_NS}}}line", {
            "x1": str(x_cursor), "y1": str(title_height),
            "x2": str(x_cursor), "y2": str(table_height),
            "stroke": "#000", "stroke-width": "0.8"
        })
        ET.SubElement(table, f"{{{SVG_NS}}}text", {
            "x": str(x_cursor + padding), "y": str(title_height + 14),
            "font-size": "8", "font-family": "Arial", "font-weight": "bold"
        }).text = str(col)[:15]
        x_cursor += col_widths[i]
    for row_index, row in df.iterrows():
        y = title_height + header_height + (row_index * row_height)
        ET.SubElement(table, f"{{{SVG_NS}}}line", {
            "x1": "0", "y1": str(y), "x2": str(table_width), "y2": str(y),
            "stroke": "#000", "stroke-width": "0.5"
        })
        x_cursor = 0
        for col_index, col in enumerate(columns):
            value = str(row[col]) if pd.notna(row[col]) else ""
            value = value[:18]
            ET.SubElement(table, f"{{{SVG_NS}}}text", {
                "x": str(x_cursor + padding), "y": str(y + 12),
                "font-size": "8", "font-family": "Arial"
            }).text = value
            x_cursor += col_widths[col_index]

# =================================================
# UPDATE SVG: point matching + table merge + optional column value at point
# =================================================
def _point_to_value_map(df, point_column, display_column):
    """Map normalized point id -> value from display_column for that row."""
    if not display_column or display_column not in df.columns:
        return {}
    pc = point_column if point_column in df.columns else df.columns[0]
    out = {}
    for _, row in df.iterrows():
        pt = str(row.get(pc, "")).strip().upper()
        if not pt:
            continue
        pt_norm = _normalize_point_id(pt)
        if not pt_norm:
            continue
        val = row[display_column]
        out[pt_norm] = "" if pd.isna(val) else str(val).strip()
    return out


def _normalize_point_id(pid):
    """Normalize for matching: remove spaces/dashes so BI 1, BI-1, BI1 all match."""
    if pid is None or (isinstance(pid, float) and pd.isna(pid)):
        return ""
    s = str(pid).strip().upper()
    s = re.sub(r"[\s\-]+", "", s)
    return s


def _point_to_left_right_map(df, point_column, left_column, right_column):
    """Maps normalized point id -> (left_value, right_value) for matching rows."""
    pc = point_column if point_column in df.columns else df.columns[0]
    left_col = left_column if left_column and left_column in df.columns else None
    right_col = right_column if right_column and right_column in df.columns else None
    if not left_col and not right_col:
        return {}
    out = {}
    for _, row in df.iterrows():
        pt = str(row.get(pc, "")).strip().upper()
        if not pt:
            continue
        pt_norm = _normalize_point_id(pt)
        if not pt_norm:
            continue
        lv = "" if not left_col or pd.isna(row.get(left_col)) else str(row[left_col]).strip()
        rv = "" if not right_col or pd.isna(row.get(right_col)) else str(row[right_col]).strip()
        out[pt_norm] = (lv, rv)
    return out


# Fixed IDs for left/right data placement (user-specified)
LEFT_DATA_ID = "data-ui1"
RIGHT_DATA_ID = "data-ui2"
# Image (e.g. id="24Vac" or "bo1-image") visible only when point matches AND SIGNAL has 24Vac


def _row_contains_24vac(row):
    """True if any cell in the row contains '24' and 'vac' (case-insensitive)."""
    for v in row.values:
        s = str(v).strip().lower() if pd.notna(v) else ""
        if "24" in s and "vac" in s:
            return True
    return False


def _signal_has_24vac(signal_val):
    """True if signal value contains '24' and 'vac' (case-insensitive)."""
    s = (str(signal_val or "").strip()).lower()
    return "24" in s and "vac" in s


def _normalize_signal_for_match(s):
    """Normalize signal string for matching with image id: lower, strip, remove spaces."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).strip().lower().replace(" ", "")


def _signal_matches_image_id(signal_val, image_id):
    """True when signal value matches the image id (e.g. '24Vac' matches '24Vac' or '24 VAC'); '0-10V' does not."""
    sig = _normalize_signal_for_match(signal_val)
    img = _normalize_signal_for_match(image_id)
    if not sig or not img:
        return False
    if sig == img or img in sig or sig in img:
        return True
    # Image id "24Vac": show when signal contains 24 and vac
    if "24" in img and "vac" in img and "24" in sig and "vac" in sig:
        return True
    # Image id like "bo1-image" or "bo2-image": treat as 24Vac symbol, show when signal has 24Vac
    if img.endswith("image") and "24" in sig and "vac" in sig:
        return True
    return False


def _get_signal_column(df):
    """Return SIGNAL column name if present (case-insensitive)."""
    for col in df.columns:
        if str(col).strip().upper() == "SIGNAL":
            return col
    return None


def _point_to_signal_map(df, point_column):
    """Maps normalized point id -> SIGNAL column value (string). Only when SIGNAL column exists."""
    pc = point_column if point_column in df.columns else df.columns[0]
    signal_col = _get_signal_column(df)
    if signal_col is None:
        return {}
    out = {}
    for _, row in df.iterrows():
        pt = str(row.get(pc, "")).strip()
        if not pt:
            continue
        pt_norm = _normalize_point_id(pt)
        if not pt_norm:
            continue
        out[pt_norm] = row.get(signal_col)
    return out


def _point_has_24vac_map(df, point_column):
    """Maps normalized point id -> True if that row contains '24 VAC' in any column."""
    pc = point_column if point_column in df.columns else df.columns[0]
    out = {}
    for _, row in df.iterrows():
        pt = str(row.get(pc, "")).strip()
        if not pt:
            continue
        pt_norm = _normalize_point_id(pt)
        if not pt_norm:
            continue
        if _row_contains_24vac(row):
            out[pt_norm] = True
    return out


def _find_by_id(container, id_val):
    """Return first element under container (self or descendant) with id attribute == id_val."""
    id_val = (id_val or "").strip()
    if not id_val:
        return None
    for el in container.iter():
        if (el.get("id") or "").strip() == id_val:
            return el
    return None


def _get_label_text_elements(g):
    """Return list of text elements in g that are labels (not - or +), sorted by x position (left to right)."""
    candidates = []
    for text_el in g.findall(".//{%s}text" % SVG_NS):
        cur = (text_el.text or "").strip()
        if cur in ("-", "+"):
            continue
        x = float(text_el.get("x", 0))
        candidates.append((x, text_el))
    candidates.sort(key=lambda x: x[0])
    return [t for _, t in candidates]


def _get_point_name_and_left_right_slots(g):
    """
    Point name = text with text-anchor="middle" (inside the box) – never touch.
    Left slots = non–point-name texts with x < point_name_x (1). Right slots = x > point_name_x (1).
    Returns (point_name_el or None, left_slots_list, right_slots_list).
    """
    texts = _get_label_text_elements(g)
    if not texts:
        return None, [], []
    middle = None
    others = []
    for t in texts:
        if (t.get("text-anchor") or "").lower() == "middle":
            middle = t
        else:
            others.append(t)
    if not others:
        return middle, [], []
    others_sorted = sorted(others, key=lambda t: float(t.get("x", 0)))
    mid_x = float(middle.get("x", 0)) if middle is not None else None
    if mid_x is not None:
        left_slots = [t for t in others_sorted if float(t.get("x", 0)) < mid_x][:1]
        right_slots = [t for t in others_sorted if float(t.get("x", 0)) > mid_x][:1]
    else:
        n = len(others_sorted)
        left_slots = others_sorted[:1]
        right_slots = others_sorted[-1:] if n > 1 and others_sorted[-1] is not others_sorted[0] else []
    return middle, left_slots, right_slots


def _set_point_label_left_right(g, left_value, right_value):
    """Set left value in id=data-ui1, right value in id=data-ui2. Fallback to position-based slots if IDs not found."""
    left_val = (left_value or "").strip()
    right_val = (right_value or "").strip()
    left_el = _find_by_id(g, LEFT_DATA_ID)
    right_el = _find_by_id(g, RIGHT_DATA_ID)
    if left_el is not None:
        left_el.text = left_val
    if right_el is not None:
        right_el.text = right_val
    if left_el is not None and right_el is not None:
        return
    _point_name_el, left_slots, right_slots = _get_point_name_and_left_right_slots(g)
    for el in left_slots:
        el.text = left_val
    for el in right_slots:
        el.text = right_val
    if not left_slots and not right_slots:
        pass
    elif not left_slots and right_slots:
        if left_val and right_val:
            right_slots[0].text = (left_val + "  " + right_val).strip()
        elif left_val:
            right_slots[0].text = left_val
    elif left_slots and not right_slots:
        if left_val and right_val:
            left_slots[0].text = (left_val + "  " + right_val).strip()
        elif right_val:
            left_slots[0].text = right_val


def _set_point_label_spare(g):
    """If point name does not match Excel: find id=data-ui1 and id=data-ui2 in this <g>, replace their text with SPARE."""
    spare = "SPARE"
    for el in g.iter():
        eid = (el.get("id") or "").strip()
        if eid != LEFT_DATA_ID and eid != RIGHT_DATA_ID:
            continue
        el.text = spare
        for child in list(el):
            el.remove(child)


def update_svg(svg_path, df, output_svg, point_column="POINT", display_column=None,
               left_column=None, right_column=None):
    tree = ET.parse(svg_path)
    root = tree.getroot()
    # Match point ID with Excel (normalized: BI 1, BI-1, BI1 all match)
    pc = point_column if point_column in df.columns else df.columns[0]
    excel_point_ids = set(_normalize_point_id(v) for v in df[pc].dropna().astype(str))
    excel_point_ids.discard("")
    point_to_lr = _point_to_left_right_map(df, point_column, left_column, right_column)
    point_to_value = _point_to_value_map(df, point_column, display_column) if display_column else {}
    # Image visible only when: point name matches AND row SIGNAL value matches the image id (e.g. "24Vac")
    point_to_signal = _point_to_signal_map(df, pc)
    point_pattern = re.compile(r"^(UI|AI|DI|AO|DO|BO|BI|NODE)[\w\-]+$", re.IGNORECASE)

    def _group_has_data_ui(grp):
        return any(
            (el.get("id") or "").strip() in (LEFT_DATA_ID, RIGHT_DATA_ID)
            for el in grp.iter()
        )

    # 1) Every <g> with id that has data-ui1/data-ui2: match id with Excel point → print column data; else replace with SPARE
    for g in root.iter():
        tag = g.tag if hasattr(g.tag, "endswith") else str(g.tag)
        if tag != "g" and not tag.endswith("}g"):
            continue
        gid = g.get("id")
        if not gid:
            continue
        if not _group_has_data_ui(g):
            continue
        # Only process leaf point groups (no nested <g> with id that also has data-ui1/data-ui2)
        has_nested = False
        for child in g:
            ctag = child.tag if hasattr(child.tag, "endswith") else str(child.tag)
            if (ctag == "g" or ctag.endswith("}g")) and child.get("id") and _group_has_data_ui(child):
                has_nested = True
                break
        if has_nested:
            continue
        gid_clean = (gid or "").strip().upper()
        gid_norm = _normalize_point_id(gid_clean)
        if gid_norm in excel_point_ids:
            if gid_norm in point_to_lr:
                left_val, right_val = point_to_lr[gid_norm]
                _set_point_label_left_right(g, left_val, right_val)
            elif gid_norm in point_to_value:
                _set_point_label_left_right(g, point_to_value[gid_norm], "")
            else:
                _set_point_label_left_right(g, "", "")
        else:
            _set_point_label_spare(g)

    # 2) Image visibility (24Vac): only for groups matching point pattern
    for g in root.findall(".//{%s}g" % SVG_NS):
        gid = g.get("id")
        if not gid:
            continue
        gid_clean = (gid or "").strip().upper()
        gid_norm = _normalize_point_id(gid_clean)
        if not point_pattern.match(gid_clean):
            continue
        point_signal = point_to_signal.get(gid_norm) if gid_norm in excel_point_ids else None
        for img_el in g.iter():
            tag = img_el.tag if hasattr(img_el.tag, "endswith") else str(img_el.tag)
            if tag != "image" and not tag.endswith("}image"):
                continue
            image_id = (img_el.get("id") or "").strip()
            show_image = (
                gid_norm in excel_point_ids
                and _signal_matches_image_id(point_signal, image_id)
            )
            img_el.set("visibility", "visible" if show_image else "hidden")
    if "viewBox" in root.attrib:
        vb = list(map(float, root.attrib["viewBox"].split()))
        width, height = vb[2], vb[3]
    else:
        width = float(root.attrib.get("width", 1600))
        height = float(root.attrib.get("height", 1000))
    new_width = width + 350
    new_height = height + 120
    root.set("viewBox", f"0 0 {new_width} {new_height}")
    root.set("width", str(new_width))
    root.set("height", str(new_height))
    drawing_group = ET.Element(f"{{{SVG_NS}}}g", {
        "transform": "translate(40,120) scale(0.7)"
    })
    for child in list(root):
        drawing_group.append(child)
    root.clear()
    root.append(drawing_group)
    append_full_excel_table(root, df)
    tree.write(output_svg, encoding="utf-8", xml_declaration=True)
    convert_to_visio_svg(output_svg, output_svg)

# =================================================
# CONVERT SVG TO VISIO-COMPATIBLE FORMAT
# =================================================
def _data_uri_svg_xml_to_png(data_uri):
    """Convert data:image/svg+xml;base64,... to data:image/png;base64,... so Visio can show it."""
    if not HAS_CAIROSVG:
        return None
    s = data_uri.strip()
    if not s.lower().startswith("data:image/svg+xml;base64,"):
        return None
    try:
        b64 = s.split(",", 1)[1]
        b64_clean = "".join(b64.split())
        svg_bytes = base64.b64decode(b64_clean)
    except Exception:
        return None
    try:
        png_bytes = cairosvg.svg2png(bytestring=svg_bytes)
        new_b64 = base64.b64encode(png_bytes).decode("ascii")
        return "data:image/png;base64," + new_b64
    except Exception:
        return None


def _parse_svg_css(style_text):
    """Parse CSS from <style> content: return dict class_name -> rule string."""
    if not style_text or not style_text.strip():
        return {}
    out = {}
    for m in re.finditer(r"\.([\w-]+)\s*\{([^{}]*)\}", style_text, re.DOTALL):
        name, body = m.group(1), m.group(2)
        out[name.strip()] = " ".join(body.split()).strip()
    return out


def _inline_css_on_elements(root, css_map):
    """Set style attribute on each element that has class= so Visio shows same as browser."""
    if not css_map:
        return
    for el in root.iter():
        cls = el.get("class")
        if not cls:
            continue
        parts = [p.strip() for p in str(cls).split() if p.strip()]
        rules = [css_map[p] for p in parts if p in css_map]
        if not rules:
            continue
        combined = "; ".join(r for r in rules if r)
        existing = (el.get("style") or "").strip()
        if existing:
            el.set("style", existing.rstrip(";") + "; " + combined)
        else:
            el.set("style", combined)


def convert_to_visio_svg(input_path, output_path):
    """
    Make SVG valid for Microsoft Visio: same drawing as browser, editable.
    Inline CSS onto elements so Visio shows styles; keep xlink, defs, images.
    """
    tree = ET.parse(input_path)
    root = tree.getroot()

    root.set("xmlns", SVG_NS)
    root.set("version", "1.1")

    # Visio expects width/height with units (px) when present
    w = root.get("width")
    h = root.get("height")
    if w and re.match(r"^\d*\.?\d+$", str(w).strip()):
        root.set("width", str(w).strip() + "px")
    if h and re.match(r"^\d*\.?\d+$", str(h).strip()):
        root.set("height", str(h).strip() + "px")

    # viewBox: keep as-is but ensure clean format (no scientific notation)
    vb = root.get("viewBox")
    if vb:
        parts = vb.replace(",", " ").split()
        if len(parts) == 4:
            try:
                nums = [float(x) for x in parts]
                root.set("viewBox", " ".join(str(int(x) if x == int(x) else x) for x in nums))
            except (ValueError, TypeError):
                pass

    for el in root.iter():
        if el.get("visibility") == "hidden":
            el.set("display", "none")
            del el.attrib["visibility"]

    # Inline CSS so Visio shows same styles as browser (Visio often ignores <style> block)
    for style_el in root.iter():
        tag = style_el.tag.split("}")[-1] if "}" in style_el.tag else style_el.tag
        if tag != "style":
            continue
        raw = style_el.text or ""
        for child in style_el:
            if child.text:
                raw += child.text
            if child.tail:
                raw += child.tail
        css_map = _parse_svg_css(raw)
        if css_map:
            _inline_css_on_elements(root, css_map)
        break

    # Ensure path/line (wires) have stroke so they print properly
    for el in root.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag not in ("path", "line"):
            continue
        if el.get("style") or el.get("class"):
            continue
        el.set("style", "stroke:#000000;stroke-width:1;fill:none")

    # Remove only unknown namespace attributes (keep w3.org, xlink, Visio)
    for el in root.iter():
        to_drop = [
            k for k in el.attrib
            if k.startswith("{")
            and "w3.org" not in k
            and "schemas.microsoft.com" not in k
        ]
        for k in to_drop:
            del el.attrib[k]

    # Declare xlink and Visio on root so styles and drawing stay correct in Visio
    has_xlink = any(
        k.startswith("{") and "1999/xlink" in k
        for el in root.iter() for k in el.attrib
    )
    if has_xlink:
        root.set("xmlns:xlink", XLINK_NS)
    has_visio = any(
        getattr(el, "tag", "").startswith("{") and "schemas.microsoft.com" in getattr(el, "tag", "")
        for el in root.iter()
    ) or any(
        k.startswith("{") and "schemas.microsoft.com" in k
        for el in root.iter() for k in el.attrib
    )
    if has_visio:
        root.set("xmlns:v", VISIO_NS)

    # For <image>: normalize data URI; convert image/svg+xml to PNG so Visio can show it
    for el in root.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag != "image":
            continue
        href_key = None
        for k in el.attrib:
            if k.startswith("{") and "1999/xlink" in k and "href" in k:
                href_key = k
                break
        if href_key is None:
            continue
        val = el.attrib[href_key]
        if not isinstance(val, str) or not val.strip().lower().startswith("data:"):
            continue
        val = "".join(val.split())
        if val.lower().startswith("data:image/svg+xml;base64,"):
            png_uri = _data_uri_svg_xml_to_png(val)
            if png_uri is not None:
                val = png_uri
        el.set("href", val)
        el.attrib[href_key] = val

    # Write with default namespace so output is clean SVG (no ns0: prefix)
    with open(output_path, "wb") as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        _serialize_visio_svg(root, f, is_root=True)


def _escape_text(s):
    if not s:
        return ""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _attr_value_for_serialize(k, v):
    """Escape attribute value; collapse whitespace in data: URIs so base64 images work in Visio."""
    s = str(v)
    if s.strip().lower().startswith("data:"):
        s = "".join(s.split())
    return _escape_text(s)


def _serialize_visio_svg(el, f, indent=0, is_root=False):
    """Write one element and children; preserve SVG, xlink, and Visio (v:) so styles and images work."""
    space = "  " * indent
    if "}" in el.tag:
        ns_uri, local = el.tag[1:].split("}", 1)
        tag = f"v:{local}" if "schemas.microsoft.com" in ns_uri else local
    else:
        tag = el.tag
    attrs = []
    for k, v in el.attrib.items():
        val_esc = _attr_value_for_serialize(k, v)
        if k.startswith("{"):
            ns_uri = k[1:].split("}", 1)[0]
            local = k.split("}")[-1]
            if "1999/xlink" in ns_uri:
                attrs.append(f' xlink:{local}="{val_esc}"')
            elif "schemas.microsoft.com" in ns_uri:
                attrs.append(f' v:{local}="{val_esc}"')
            elif "w3.org" in ns_uri and "XML" in ns_uri:
                attrs.append(f' xml:{local}="{val_esc}"')
            continue
        if is_root and tag == "svg" and k in ("xmlns", "version"):
            continue
        attrs.append(f' {k}="{val_esc}"')
    if is_root and tag == "svg":
        attrs.insert(0, f' xmlns="{SVG_NS}"')
        attrs.append(' version="1.1"')
        # xmlns:xlink and xmlns:v are already on root.attrib from convert_to_visio_svg
    attr_str = "".join(attrs)
    has_children = len(el) > 0
    text = (el.text or "").strip()
    # Style/script: output content as CDATA so CSS and special chars are preserved for Visio
    is_style_or_script = tag in ("style", "script")
    if not has_children and not text and not is_style_or_script:
        f.write((space + f"<{tag}{attr_str}/>\n").encode("utf-8"))
        return
    f.write((space + f"<{tag}{attr_str}>").encode("utf-8"))
    if is_style_or_script and (text or (has_children and el.text)):
        content = el.text or ""
        for c in el:
            if c.text:
                content += c.text
            if c.tail:
                content += c.tail
        if content.strip():
            f.write(b"\n<![CDATA[\n")
            f.write(content.encode("utf-8"))
            f.write(b"\n]]>\n")
        f.write((space + f"</{tag}>\n").encode("utf-8"))
        return
    if text and not has_children:
        f.write(_escape_text(el.text).encode("utf-8"))
        f.write((f"</{tag}>\n").encode("utf-8"))
        return
    if text:
        f.write(("\n" + "  " * (indent + 1) + _escape_text(el.text)).encode("utf-8"))
    f.write(b"\n")
    for i, child in enumerate(el):
        _serialize_visio_svg(child, f, indent + 1, is_root=False)
        if child.tail and child.tail.strip():
            f.write(("  " * (indent + 1) + _escape_text(child.tail.strip()) + "\n").encode("utf-8"))
    f.write((space + f"</{tag}>\n").encode("utf-8"))
@app.after_request
def _no_cache_html(response):
    """Avoid cached pages on other PC so the app always loads fresh (no reload loop)."""
    if response.content_type and "text/html" in response.content_type:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
    return response


@app.route("/health")
def health():
    """Health check for hosting platforms (Render, Railway, etc.). Returns 200 when app is up."""
    return "ok", 200


@app.route("/favicon.ico")
def favicon():
    """Avoid 404 when browser requests favicon."""
    return "", 204


@app.route("/step1")
def step1_redirect():
    """Redirect /step1 to / so old links and bookmarks work."""
    return redirect(url_for("step1"), code=302)


@app.route("/", methods=["GET", "POST"])
def step1():
    if request.method == "POST":
        excel = request.files.get("excel")
        if not excel or not excel.filename or not excel.filename.lower().endswith((".xlsx", ".xls")):
            return redirect(url_for("step1") + "?error=upload"), 302
        try:
            sheet = int(request.form.get("sheet", 0))
        except (ValueError, TypeError):
            sheet = 0
        path = os.path.join(EXCEL_DIR, excel.filename)
        try:
            excel.save(path)
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception:
            return redirect(url_for("step1") + "?error=excel"), 302
        tables = read_all_tables(raw)
        table_ids = []

        for tdf in tables:
            prefix = safe_col(tdf, 0).astype(str).str.strip()
            numbers = pd.to_numeric(safe_col(tdf, 1), errors="coerce")

            clean_numbers = numbers.apply(
                lambda x: str(int(x)) if pd.notna(x) and float(x).is_integer()
                else (str(x) if pd.notna(x) else "")
            )

            point_col = prefix + clean_numbers

            df = pd.DataFrame({
                "POINT": point_col,
                "SYSTEM": safe_col(tdf, 2).astype(str),
                "OBJECT": safe_col(tdf, 3).astype(str),
                "DESCRIPTION": safe_col(tdf, 4).astype(str),
                "SIGNAL": safe_col(tdf, 5).astype(str),
            }).fillna("").replace("nan", "")

            tid = str(uuid.uuid4())

            # Save SVG
            svg = build_table_svg(df)
            svg_path = os.path.join(TEMP_DIR, f"{tid}.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg)

            # Save Excel
            excel_out = os.path.join(TEMP_DIR, f"{tid}.xlsx")
            df.to_excel(excel_out, index=False)

            table_ids.append(tid)

        session["table_ids"] = table_ids
        return render_template("step1.html", table_ids=table_ids, preview_refresh="")

    preview_refresh = request.args.get("refresh")
    upload_error = request.args.get("error")
    table_ids = session.get("table_ids") or []
    return render_template("step1.html", table_ids=table_ids, preview_refresh=preview_refresh, upload_error=upload_error)

# =================================================
# PREVIEW SVG
# =================================================
@app.route("/preview/<pid>")
def preview(pid):
    path = os.path.join(TEMP_DIR, f"{pid}.svg")
    if not os.path.isfile(path):
        return "Not found", 404
    resp = send_file(path)
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp

# =================================================
# DOWNLOAD GENERATED EXCEL
# =================================================
@app.route("/download_excel/<pid>")
def download_excel(pid):
    return send_file(os.path.join(TEMP_DIR, f"{pid}.xlsx"), as_attachment=True)


# =================================================
# EDIT TABLE – change values before generating
# =================================================
@app.route("/edit-table/<table_id>", methods=["GET", "POST"])
def edit_table(table_id):
    excel_path = os.path.join(TEMP_DIR, f"{table_id}.xlsx")
    if not os.path.isfile(excel_path):
        return "Table not found. Go back and create tables first.", 404
    if request.method == "POST":
        # Build DataFrame from form: data_0_POINT, data_0_SYSTEM, ... data_1_POINT, ...
        columns = ["POINT", "SYSTEM", "OBJECT", "DESCRIPTION", "SIGNAL"]
        rows = []
        row_index = 0
        while True:
            key = f"data_{row_index}_POINT"
            if key not in request.form:
                break
            row = {}
            for col in columns:
                val = request.form.get(f"data_{row_index}_{col}", "")
                row[col] = str(val).strip() if val else ""
            rows.append(row)
            row_index += 1
        if not rows:
            return "No data submitted.", 400
        df = pd.DataFrame(rows, columns=columns)
        df.to_excel(excel_path, index=False)
        # Regenerate table SVG so preview stays in sync
        svg = build_table_svg(df)
        svg_path = os.path.join(TEMP_DIR, f"{table_id}.svg")
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(svg)
        # Redirect with refresh so step1 loads new preview (no cache)
        import time
        return redirect(url_for("step1", refresh=int(time.time() * 1000)))
    df = pd.read_excel(excel_path)
    cols = [str(c).strip() for c in df.columns]
    want = ["POINT", "SYSTEM", "OBJECT", "DESCRIPTION", "SIGNAL"]
    for c in want:
        if c not in cols:
            df[c] = ""
    # Use only wanted columns in order
    df = df.reindex(columns=want).fillna("")
    df = df.astype(str)
    return render_template("edit_table.html", table_id=table_id, df=df, table_index=request.args.get("table_index", "1"))


# =================================================
# MERGE TABLE INTO DRAWING
# =================================================
@app.route("/merge", methods=["POST"])
def merge():
    drawing = request.files["drawing"]
    table_id = request.form["table_id"]

    drawing_path = os.path.join(DRAWING_DIR, drawing.filename)
    drawing.save(drawing_path)

    drawing_tree = ET.parse(drawing_path)
    drawing_root = drawing_tree.getroot()
    ET.register_namespace("", SVG_NS)

    draw_w, draw_h = get_viewbox(drawing_root)

    table_tree = ET.parse(os.path.join(TEMP_DIR, f"{table_id}.svg"))
    table_root = table_tree.getroot()

    table_w = float(table_root.attrib["width"])
    table_h = float(table_root.attrib["height"])

    scale = min((draw_w * 0.45) / table_w, (draw_h * 0.7) / table_h, 1)
    tx = draw_w - (table_w * scale) - 20
    ty = 40

    g = ET.Element(svg_tag("g"), {
        "transform": f"translate({tx},{ty}) scale({scale})"
    })

    for el in list(table_root):
        g.append(el)

    drawing_root.append(g)

    out = os.path.join(TEMP_DIR, f"FINAL_{uuid.uuid4()}.svg")
    drawing_tree.write(out, encoding="utf-8", xml_declaration=True)

    return send_file(out, as_attachment=True)

# =================================================
# MERGE DASHBOARD – select table or upload Excel + SVG
# =================================================
@app.route("/merge-dashboard", methods=["GET"])
def merge_dashboard():
    table_ids = session.get("table_ids") or []
    svg_templates = list_svg_templates()
    return render_template("merge_dashboard.html", table_ids=table_ids, svg_templates=svg_templates)


# =================================================
# SAVE SVG AS TEMPLATE
# =================================================
@app.route("/save-svg-template", methods=["POST"])
def save_svg_template():
    name = (request.form.get("template_name") or "").strip()
    svg_file = request.files.get("svg_file")
    if not name:
        return "Template name is required.", 400
    if not svg_file or not svg_file.filename or not svg_file.filename.lower().endswith("output.svg"):
        return "Please upload an SVG file.", 400
    safe_name = re.sub(r"[^\w\s-]", "", name).strip().replace(" ", "_") or "template"
    filename = f"{safe_name}.svg"
    path = os.path.join(SVG_TEMPLATES_DIR, filename)
    svg_file.save(path)
    return redirect(url_for("merge_dashboard"))

# =================================================
# MERGE FINAL – same process: point match + table in drawing
# =================================================
@app.route("/merge-final", methods=["POST"])
def merge_final():
    table_source = request.form.get("table_source", "selected")
    svg_source = request.form.get("svg_source", "upload")
    svg_file = request.files.get("svg_file")
    svg_template = request.form.get("svg_template")

    if svg_source == "template":
        if not svg_template:
            return "Please select a saved SVG template.", 400
        path = os.path.join(SVG_TEMPLATES_DIR, svg_template)
        if not os.path.isfile(path):
            return "Selected SVG template not found.", 404
        svg_path = path
    else:
        if not svg_file or not svg_file.filename:
            return "Please upload an SVG drawing or select a saved template.", 400
        svg_path = os.path.join(TEMP_DIR, f"input_drawing_{uuid.uuid4()}.svg")
        svg_file.save(svg_path)

    if table_source == "upload":
        excel_file = request.files.get("excel_file")
        if not excel_file:
            return "Please upload an Excel file when choosing 'Upload new Excel'.", 400
        excel_path = os.path.join(TEMP_DIR, f"input_excel_{uuid.uuid4()}.xlsx")
        excel_file.save(excel_path)
        df = pd.read_excel(excel_path)
    else:
        table_id = request.form.get("table_id")
        if not table_id:
            return "Please select a table.", 400
        excel_path = os.path.join(TEMP_DIR, f"{table_id}.xlsx")
        if not os.path.isfile(excel_path):
            return "Selected table file not found. Go back and create tables first.", 404
        df = pd.read_excel(excel_path)

    point_column = (request.form.get("point_column") or "POINT").strip() or "POINT"
    display_column = (request.form.get("display_column") or "").strip() or None
    left_column = (request.form.get("left_column") or "").strip() or None
    right_column = (request.form.get("right_column") or "").strip() or None

    # User-provided download filename (optional)
    raw_name = (request.form.get("output_filename") or "").strip()
    if raw_name:
        safe = re.sub(r"[^\w\s\-.]", "", raw_name).strip().replace(" ", "_") or "final_output"
        download_name = safe if safe.lower().endswith(".svg") else (safe + ".svg")
    else:
        download_name = "final_output.svg"

    output_svg = os.path.join(TEMP_DIR, f"final_output_{uuid.uuid4()}.svg")
    update_svg(svg_path, df, output_svg, point_column=point_column, display_column=display_column,
              left_column=left_column, right_column=right_column)
    return send_file(output_svg, as_attachment=True, download_name=download_name)

# =================================================
# FINAL RUNNING LINK (for other PCs)
# =================================================
# Port: change PORT.txt in this folder, or set env PORT=8080, or run: python app.py --port 8080
def _default_port():
    try:
        if os.environ.get("PORT"):
            return int(os.environ.get("PORT", "").strip())
    except (ValueError, TypeError):
        pass
    port_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PORT.txt")
    if os.path.isfile(port_file):
        try:
            with open(port_file, "r", encoding="utf-8") as f:
                line = f.readline().strip()
                if line:
                    return int(line)
        except (ValueError, TypeError, OSError):
            pass
    return 6001


PORT = _default_port()
# Set when running with --host-public and pyngrok; public URL for any PC
PUBLIC_LINK = None


def _local_ips():
    """Get local IPs so other PCs can connect (skip 127.x)."""
    out = []
    seen = set()
    try:
        import socket
        for _ in ("", "127.0.0.1"):
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                s.settimeout(0.5)
                s.connect(("8.8.8.8", 80))
                ip = s.getsockname()[0]
                s.close()
                if ip and not ip.startswith("127.") and ip not in seen:
                    out.append(ip)
                    seen.add(ip)
            except Exception:
                pass
        if not out:
            try:
                host = socket.gethostbyname(socket.gethostname())
                if host and not host.startswith("127.") and host not in seen:
                    out.append(host)
                    seen.add(host)
            except Exception:
                pass
        # Windows: get all IPv4 from ipconfig so we don't miss LAN IP
        try:
            import subprocess
            r = subprocess.run(
                ["ipconfig"], capture_output=True, text=True, timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
            )
            if r.returncode == 0 and r.stdout:
                for line in r.stdout.splitlines():
                    line = line.strip()
                    if "IPv4" in line or "IP Address" in line:
                        # e.g. "IPv4 Address. . . . . . . . . . . : 192.168.1.5"
                        parts = line.split(":")
                        if len(parts) >= 2:
                            ip = parts[-1].strip()
                            if ip and ip not in seen and not ip.startswith("127."):
                                try:
                                    socket.inet_pton(socket.AF_INET, ip)
                                    out.append(ip)
                                    seen.add(ip)
                                except Exception:
                                    pass
        except Exception:
            pass
    except Exception:
        pass
    return out


def _is_hosted_online():
    """True if app is reached via a non-local URL (deployed online)."""
    try:
        root = (request.url_root or "").strip().lower()
        return root and "localhost" not in root and "127.0.0.1" not in root
    except Exception:
        return False


@app.route("/open-on-other-pc")
def open_on_other_pc_help():
    """Help page: why the link doesn't open on another PC and how to fix it."""
    ips = _local_ips()
    port = PORT
    links = ["http://%s:%s" % (ip, port) for ip in ips] if ips else []
    return render_template("open_on_other_pc.html", links=links, port=port)


@app.route("/link")
def running_link():
    """Page that shows the app URL. When hosted online: only the app URL. When local: local + other PC links."""
    app_url = (request.url_root or "").rstrip("/")
    hosted_online = _is_hosted_online()
    if hosted_online:
        return render_template(
            "running_link.html",
            hosted_online=True,
            link_app_url=app_url,
            link_local=app_url,
            link_other=None,
            links_other=[],
            link_public=app_url,
            port=request.environ.get("SERVER_PORT", PORT),
        )
    ips = _local_ips()
    link_local = "http://localhost:%s" % PORT
    link_other = ("http://%s:%s" % (ips[0], PORT)) if ips else None
    links_other = ["http://%s:%s" % (ip, PORT) for ip in ips] if ips else []
    return render_template(
        "running_link.html",
        hosted_online=False,
        link_app_url=None,
        link_local=link_local,
        link_other=link_other,
        links_other=links_other,
        link_public=PUBLIC_LINK,
        port=PORT,
    )


if __name__ == "__main__":
    import sys
    import time
    import logging
    from threading import Thread

    # Suppress "development server" warning in console (not an error)
    logging.getLogger("werkzeug").setLevel(logging.ERROR)

    port = PORT
    # Allow override: python app.py --port 8080
    if "--port" in sys.argv:
        idx = sys.argv.index("--port")
        if idx + 1 < len(sys.argv):
            try:
                port = int(sys.argv[idx + 1])
                globals()["PORT"] = port
            except ValueError:
                pass
    if os.environ.get("PORT"):
        try:
            port = int(os.environ.get("PORT", "").strip())
            globals()["PORT"] = port
        except (ValueError, TypeError):
            pass

    host_public = "--host-public" in sys.argv or os.environ.get("HOST_PUBLIC", "").strip().lower() in ("1", "true", "yes")

    # Run without debug so other PC gets a stable connection (no reload/drop). Set FLASK_DEBUG=1 to enable debug.
    is_production = bool(os.environ.get("PRODUCTION") or os.environ.get("RENDER") or os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("FLASK_ENV") == "production")
    debug = is_production is False and os.environ.get("FLASK_DEBUG", "").strip() in ("1", "true", "yes")

    if host_public:
        # Start Flask in background so we can attach ngrok and get a public link
        def run_flask():
            app.run(host="0.0.0.0", port=port, debug=debug, use_reloader=False)

        print("\n  Starting server for public link (host)...")
        flask_thread = Thread(target=run_flask, daemon=False)
        flask_thread.start()
        # Wait for Flask to be ready before opening tunnel
        time.sleep(4)

        try:
            from pyngrok import ngrok
            authtoken = os.environ.get("NGROK_AUTHTOKEN", "").strip()
            if not authtoken:
                print("  NGROK_AUTHTOKEN not set. Get a free token at https://ngrok.com and run:")
                print("    set NGROK_AUTHTOKEN=your_token")
                print("  Then run: python app.py --host-public")
            else:
                ngrok.set_auth_token(authtoken)
                # addr = local port to forward to; proto = http
                tunnel = ngrok.connect(addr=str(port), proto="http")
                public_url = tunnel.public_url if hasattr(tunnel, "public_url") else str(tunnel)
                globals()["PUBLIC_LINK"] = public_url
                print("\n" + "=" * 64)
                print("  BMS Point Tool - HOSTED (open from any PC)")
                print("=" * 64)
                print("  Public link (share with any PC):  %s" % public_url)
                print("  This PC:                          http://localhost:%s" % port)
                print("-" * 64)
                print("  On another PC: open the link above. If you see an ngrok")
                print("  warning page, click 'Visit Site' to reach the app.")
                print("=" * 64)
                print("  Or open %s/link to copy the public link." % public_url)
                print("=" * 64 + "\n")
        except Exception as e:
            print("  Could not start public link: %s" % e)
            print("  Run without --host-public for local/same-network only.")
        flask_thread.join()

    else:
        ips = _local_ips()
        link_local = "http://localhost:%s" % port
        print("\n" + "=" * 60)
        print("  BMS Point Tool - Server running")
        print("=" * 60)
        print("  This PC:            %s" % link_local)
        if ips:
            print("  Other PCs - copy one of these (same WiFi/LAN):")
            for ip in ips:
                print("    http://%s:%s" % (ip, port))
            print("  >>> If other PC only reloads/does not open: run allow_firewall.bat as Admin, or")
            print("  >>> use Host for other PCs.bat (ngrok) / deploy to Render - see DEPLOY.md")
        else:
            print("  Other PCs:  http://<THIS_PC_IP>:%s  (run ipconfig to get IP)" % port)
            print("  >>> If other PC cannot open: run allow_firewall.bat as Administrator.")
        print("=" * 60)
        print("  Listening on 0.0.0.0:%s (no debug = stable for other PC)" % port)
        print("=" * 60 + "\n")
        app.run(host="0.0.0.0", port=port, debug=debug, use_reloader=False)