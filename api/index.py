from flask import Flask, request, send_file, jsonify
import io
import os

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({"message": "API is running", "endpoint": "/generate"})

@app.route('/generate', methods=['POST'])
def generate():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file"}), 400
            
        file = request.files['file']
        
        # Import here to save memory
        import pandas as pd
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
        from pptx.enum.text import PP_ALIGN
        from collections import defaultdict, namedtuple
        
        df = pd.read_excel(io.BytesIO(file.read())).fillna("")
        
        # Simple validation
        required = {"Name", "Title", "Department", "Reports to", "Flag"}
        if not required.issubset(df.columns):
            missing = required - set(df.columns)
            return jsonify({"error": f"Missing columns: {missing}"}), 400
        
        # Create simple PPTX for testing
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Add title
        title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tf = title.text_frame
        p = tf.paragraphs[0]
        p.text = "Org Chart - Test"
        p.font.size = Pt(24)
        p.font.bold = True
        
        # Add some data
        y_pos = Inches(2)
        for _, row in df.head(10).iterrows():
            box = slide.shapes.add_textbox(Inches(1), y_pos, Inches(8), Inches(0.5))
            tf = box.text_frame
            p = tf.paragraphs[0]
            p.text = f"{row.get('Name', '')} - {row.get('Title', '')}"
            p.font.size = Pt(12)
            y_pos += Inches(0.6)
        
        out = io.BytesIO()
        prs.save(out)
        
        return send_file(
            io.BytesIO(out.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='test.pptx'
        )
        
    except Exception as e:
        import traceback
        return jsonify({
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

if __name__ == '__main__':
    app.run(debug=True)
