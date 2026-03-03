from flask import Flask, request, send_file, jsonify
import io
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from collections import defaultdict, namedtuple

app = Flask(__name__)

# ---------- Styling ----------
COL_NAVY = RGBColor(11, 29, 89)
COL_GREY = RGBColor(74, 74, 74)
COL_GREEN = RGBColor(46, 158, 46)
COL_YELLOW = RGBColor(244, 194, 13)
COL_BLUE = RGBColor(30, 78, 216)
COL_WHITE = RGBColor(255, 255, 255)
COL_BLACK = RGBColor(0, 0, 0)


def _text_color(fill_rgb):
    return COL_BLACK if fill_rgb == COL_YELLOW else COL_WHITE


def _node_fill(flag: int, label: str):
    if flag == 1:
        return COL_NAVY
    lab = (label or "").lower()
    if "strong" in lab:
        return COL_GREEN
    if "possible" in lab:
        return COL_YELLOW
    if "offshore" in lab:
        return COL_BLUE
    return COL_BLUE


def _add_box(slide, x, y, w, h, title, subtitle="", footer="", fill_rgb=COL_BLUE):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_rgb
    sh.line.color.rgb = RGBColor(230, 230, 230)

    tf = sh.text_frame
    tf.clear()
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(4)
    tf.margin_bottom = Pt(4)

    tc = _text_color(fill_rgb)

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = tc

    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(10)
        r2.font.color.rgb = tc

    if footer:
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = footer
        r3.font.size = Pt(10)
        r3.font.color.rgb = tc

    return sh


def _add_connector(slide, parent_sh, child_sh):
    x1 = parent_sh.left + parent_sh.width // 2
    y1 = parent_sh.top + parent_sh.height
    x2 = child_sh.left + child_sh.width // 2
    y2 = child_sh.top
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.line.color.rgb = RGBColor(120, 120, 120)
    conn.line.width = Pt(1.25)
    return conn


def build_pptx_from_excel_bytes(xlsx_bytes: bytes) -> bytes:
    df = pd.read_excel(io.BytesIO(xlsx_bytes)).fillna("")

    required = {"Name", "Title", "Department", "Reports to", "Flag"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")

    # Root detection
    root_row = df[df["Reports to"].astype(str).str.strip() == ""]
    if len(root_row) != 1:
        root_row = df[df["Title"].astype(str).str.strip().str.upper() == "CEO"]
    if len(root_row) != 1:
        raise ValueError("Could not determine a single root (CEO). Ensure one row has blank 'Reports to'.")

    root = root_row.iloc[0]
    CEO_NAME = str(root["Name"]).strip()
    CEO_TITLE = str(root["Title"]).strip()

    Node = namedtuple("Node", "id title name dept manager label flag count")
    nodes = {}
    children = defaultdict(list)
    name_to_ids = defaultdict(list)

    next_id = 1

    def add_node(title, name, dept, manager, label, flag, count=1):
        nonlocal next_id
        nid = f"n{next_id}"
        next_id += 1
        nodes[nid] = Node(nid, title, name, dept, manager, label, flag, count)
        return nid

    # Named nodes
    for _, r in df.iterrows():
        if int(r["Flag"]) == 1:
            nid = add_node(
                str(r["Title"]).strip(),
                str(r["Name"]).strip(),
                str(r["Department"]).strip(),
                str(r["Reports to"]).strip(),
                str(r.get("Label", "")).strip(),
                1,
                1
            )
            name_to_ids[str(r["Name"]).strip()].append(nid)

    # Aggregated nodes for Flag=0
    agg = defaultdict(int)
    for _, r in df.iterrows():
        if int(r["Flag"]) == 0:
            key = (
                str(r["Title"]).strip(),
                str(r["Department"]).strip(),
                str(r["Reports to"]).strip(),
                str(r.get("Label", "")).strip()
            )
            agg[key] += 1

    for (title, dept, manager, label), cnt in agg.items():
        add_node(title, "", dept, manager, label, 0, cnt)

    # Build children mapping
    for nid, n in nodes.items():
        mgr = n.manager.strip()
        if not mgr:
            continue
        for mid in name_to_ids.get(mgr, []):
            children[mid].append(nid)

    if CEO_NAME not in name_to_ids:
        raise ValueError("CEO name not found in Flag=1 nodes. Ensure CEO row has Flag=1.")

    CEO_ID = name_to_ids[CEO_NAME][0]

    # Departments from CEO direct reports
    dept_roots = defaultdict(list)
    for cid in children.get(CEO_ID, []):
        dept_roots[nodes[cid].dept].append(cid)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    W = prs.slide_width
    margin_x = Inches(0.4)
    top_y = Inches(0.4)

    ceo_w = Inches(3.4)
    ceo_h = Inches(0.75)
    ceo_x = (W - ceo_w) // 2
    ceo_y = top_y

    ceo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ceo_x, ceo_y, ceo_w, ceo_h)
    ceo.fill.solid()
    ceo.fill.fore_color.rgb = COL_NAVY
    ceo.line.color.rgb = COL_NAVY

    tf = ceo.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = CEO_TITLE
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = COL_WHITE
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = CEO_NAME
    r2.font.size = Pt(11)
    r2.font.color.rgb = COL_WHITE

    # Column order (optional preferred)
    preferred = ["Technology", "Executive", "Sales and Marketing", "Human Resource", "Finance", "Talent Acquisition"]
    depts = [d for d in preferred if d in dept_roots]
    for d in sorted(dept_roots.keys()):
        if d not in depts:
            depts.append(d)

    ncols = max(1, len(depts))
    usable_w = W - 2 * margin_x
    col_w = usable_w // ncols

    dept_y = ceo_y + ceo_h + Inches(0.35)
    dept_header_h = Inches(0.5)
    role_h = Inches(0.65)
    role_w = Inches(2.6)
    gap_x = Inches(0.15)
    gap_y = Inches(0.35)

    for i, dept in enumerate(depts):
        col_left = margin_x + col_w * i
        header_w = col_w - Inches(0.15)
        header_x = col_left + (col_w - header_w) // 2

        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, header_x, dept_y, header_w, dept_header_h)
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = COL_GREY
        hdr.line.color.rgb = COL_GREY
        tfh = hdr.text_frame
        tfh.clear()
        ph = tfh.paragraphs[0]
        ph.alignment = PP_ALIGN.CENTER
        rh = ph.add_run()
        rh.text = dept
        rh.font.size = Pt(14)
        rh.font.bold = True
        rh.font.color.rgb = COL_WHITE

        roots = dept_roots.get(dept, [])
        base_y = dept_y + dept_header_h + Inches(0.25)

        for idx, rid in enumerate(roots):
            n = nodes[rid]
            bx = header_x + (header_w - role_w) // 2
            by = base_y + idx * (role_h + Inches(0.15))

            fill = _node_fill(n.flag, n.label)
            sh = _add_box(slide, bx, by, role_w, role_h, n.title, n.name, "", fill)
            _add_connector(slide, ceo, sh)

            kids = children.get(rid, [])
            if not kids:
                continue

            level2_y = by + role_h + gap_y
            k = len(kids)
            box_w2 = min(role_w, max(Inches(1.9), header_w / max(1, k) - gap_x))
            total_w = k * box_w2 + (k - 1) * gap_x
            start_x = header_x + (header_w - total_w) // 2

            for j, cid in enumerate(kids):
                cn = nodes[cid]
                cx = start_x + j * (box_w2 + gap_x)
                cy = level2_y
                fill2 = _node_fill(cn.flag, cn.label)
                subtitle = cn.name if cn.flag == 1 else ""
                footer = "" if cn.flag == 1 else str(cn.count)

                sh2 = _add_box(slide, cx, cy, box_w2, role_h, cn.title, subtitle, footer, fill2)
                _add_connector(slide, sh, sh2)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


@app.route('/')
def home():
    return jsonify({
        "message": "Org Chart PPTX Generator API",
        "endpoints": {
            "POST /generate": "Upload Excel file, returns PPTX",
            "GET /": "This info"
        },
        "usage": "Send POST request to /generate with file in 'file' field"
    })


@app.route('/generate', methods=['POST'])
def generate():
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({"error": "No file provided. Use 'file' field in form-data."}), 400
        
        file = request.files['file']
        
        # Check if file is selected
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        # Validate file extension
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400
        
        # Read file bytes
        xlsx_bytes = file.read()
        
        # Generate PPTX
        pptx_bytes = build_pptx_from_excel_bytes(xlsx_bytes)
        
        # Return as downloadable file
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='org-chart-editable.pptx'
        )
        
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


# Vercel serverless handler
def handler(request, **kwargs):
    return app(request.environ, lambda status, headers: None)
