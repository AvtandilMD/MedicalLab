import webbrowser
import os
import sys
import threading
import json
from flask import Flask, render_template, request, send_file, Response, jsonify
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_template_folder():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'templates')
    return os.path.join(get_base_path(), 'templates')


def get_saved_docs_folder():
    folder = os.path.join(get_base_path(), 'saved_docs')
    os.makedirs(folder, exist_ok=True)
    return folder


def get_database_path():
    return os.path.join(get_base_path(), 'patients_db.json')


app = Flask(__name__, template_folder=get_template_folder())


# ========== პაციენტების ბაზა ==========
def load_database():
    db_path = get_database_path()
    if os.path.exists(db_path):
        with open(db_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"patients": []}


def save_database(data):
    db_path = get_database_path()
    with open(db_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_patient_record(first_name, last_name, age, test_type, filename, test_date):
    db = load_database()
    record = {
        "id": len(db["patients"]) + 1,
        "first_name": first_name,
        "last_name": last_name,
        "age": age,
        "test_type": test_type,
        "filename": filename,
        "test_date": test_date,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    db["patients"].append(record)
    save_database(db)


# ========== შაბლონები ==========
CBC_TEMPLATE = {
    "cbc_analysis": [
        {"abbr": "WBC", "parameter": "ლეიკოციტი", "reference_range": "მ. 5.0-10.0; ქ. 5.0-10.0", "unit": "10^9/L"},
        {"abbr": "RBC", "parameter": "ერითროციტი", "reference_range": "მ. 4.5-5.5; ქ. 4.5-5.5", "unit": "10^12/L"},
        {"abbr": "HGB", "parameter": "ჰემოგლობინი", "reference_range": "მ. 140-174; ქ. 120-174", "unit": "g/L"},
        {"abbr": "HCT", "parameter": "ჰემატოკრიტი", "reference_range": "მ. 36-52; ქ. 45-52", "unit": "%"},
        {"abbr": "PLT", "parameter": "თრომბოციტი", "reference_range": "მ. 150-400; ქ. 150-400", "unit": "10^9/L"},
        {"abbr": "RET", "parameter": "რეტიკულოციტი", "reference_range": "მ. 2-10; ქ. 2-10", "unit": "%"},
        {"abbr": "MCV", "parameter": "ერითროც. საშუალო მოცულობა", "reference_range": "მ. 84-96; ქ. 76-96",
         "unit": "FL"},
        {"abbr": "MCH", "parameter": "HGB საშუალო შემცველობა", "reference_range": "მ. 27-32; ქ. 27-32", "unit": "pg"},
        {"abbr": "MCHC", "parameter": "HGB საშუალო კონცენტრაცია", "reference_range": "მ. 300-350; ქ. 300-350",
         "unit": "g/l"},
        {"abbr": "RDW", "parameter": "ერითროც. განაწილების ფართი", "reference_range": "მ. 20-42; ქ. 20-42",
         "unit": "%"},
        {"abbr": "MPV", "parameter": "თრომბოც. საშუალო მოცულობა", "reference_range": "მ. 8-15; ქ. 8-15", "unit": "FL"},
        {"abbr": "PDW", "parameter": "თრომბოც. განაწილების ფართი", "reference_range": "მ. - ; ქ. -", "unit": "%"},
        {"abbr": "ESR", "parameter": "ერითროც. დალექვის სიჩქარე", "reference_range": "მ. 2-10; ქ. 2-15",
         "unit": "მმ/სთ"}
    ],
    "leukocyte_formula": [
        {"parameter": "მიელოციტი (MIEL %)", "norm": "0%"}, {"parameter": "მეტამიელოციტი (METAM %)", "norm": "0%"},
        {"parameter": "ჩხირბირთვიანი ნეიტროფილი (Rod NEUT %)", "norm": "0-6%"},
        {"parameter": "სეგმენტბირთვიანი ნეიტროფილი (SEG %)", "norm": "47-72%"},
        {"parameter": "ეოზინოფილი (EO %)", "norm": "0.5-5%"}, {"parameter": "ბაზოფილი (BASO %)", "norm": "0-1%"},
        {"parameter": "ლიმფოციტი (LYMPH %)", "norm": "19-37%"}, {"parameter": "მონოციტი (MONO %)", "norm": "3-11%"},
        {"parameter": "პლაზმური უჯრედი (PLAZ %)", "norm": "0.5-1%"}
    ]
}

URINE_TEMPLATE = {
    "header": {"subtitle": "საოჯახო მედიცინის ცენტრი", "phones": ["558-27-55-51", "577-03-97-70"]},
    "test_info": {"code": "UR.7", "name": "შარდის საერთო ანალიზი"},
    "physico_chemical": [
        {"abbr": "", "parameter": "რაოდენობა", "norm": "", "unit": "მლ"},
        {"abbr": "", "parameter": "ფერი", "norm": "ჩალისფერი", "unit": ""},
        {"abbr": "", "parameter": "გამჭვირვალობა", "norm": "გამჭვირვალე", "unit": ""},
        {"abbr": "SG", "parameter": "ხვედრითი წონა", "norm": "1.005-1.030", "unit": ""},
        {"abbr": "PH", "parameter": "რეაქცია", "norm": "5.0-8.0", "unit": ""},
        {"abbr": "PRO", "parameter": "ცილა", "norm": "0", "unit": "g/l"},
        {"abbr": "GLU", "parameter": "გლუკოზა", "norm": "0", "unit": "mmol/l"},
        {"abbr": "KET", "parameter": "კეტონები", "norm": "0", "unit": "mmol/l"},
        {"abbr": "UBG", "parameter": "ურობილინოგენი", "norm": "3.4-17.0", "unit": "µmol/l"},
        {"abbr": "BIL", "parameter": "ბილირუბინი", "norm": "0", "unit": "µmol/l"},
        {"abbr": "NIT", "parameter": "ნიტრატები", "norm": "NEG", "unit": ""},
        {"abbr": "LEU", "parameter": "ლეიკოციტები", "norm": "-", "unit": "Leu/µL"},
        {"abbr": "BLD", "parameter": "ერითროციტები", "norm": "-", "unit": "Ery/µL"}
    ],
    "microscopy": {
        "epithelium": [{"key": "squamous", "label": "ბრტყელი"}, {"key": "transitional", "label": "გარდამავალი"},
                       {"key": "renal", "label": "თირკმლის"}],
        "cylinders": [{"key": "hyaline", "label": "ჰიალინური"}, {"key": "granular", "label": "მარცვლოვანი"},
                      {"key": "waxy", "label": "ცვილისებური"}],
        "others": [{"key": "mucus", "parameter": "ლორწო"}, {"key": "salts", "parameter": "მარილები"},
                   {"key": "bacteria", "parameter": "ბაქტერიები"}, {"key": "fungi", "parameter": "სოკო"}]
    },
    "footer": {"equipment": "SIEMENS CLINITEK Status+"}
}

CRP_TEMPLATE = {
    "clinic_info": {"description": "საოჯახო მედიცინის ცენტრი", "phones": ["558-27-55-51", "577-03-97-70"]},
    "test_details": {"title_ge": "მაღალი მგრძნობელობის C-რეაქტიული ცილა (BL.7.9.1)"},
    "test_results": [
        {"code": "CRP", "parameter": "C-რეაქტიული ცილა", "reference_range": "0-10", "unit": "mg/L (მგ/ლ)"},
        {"code": "hsCRP", "parameter": "მაღალი მგრძნობელობის C-რეაქტიული ცილა", "reference_range": "0-1",
         "unit": "mg/L (მგ/ლ)"}
    ]
}

TROPONIN_TEMPLATE = {
    "document_info": {
        "clinic_name": "პრემიუმ მედი",
        "clinic_description": "საოჯახო მედიცინის ცენტრი",
        "contact": "ტელ: 558-27-55-51, 577-03-97-70"
    },
    "test_info": {
        "title": "ტროპონინის ტესტი (BL.7.8)",
        "results_table": [
            {
                "code": "BL.7.8",
                "parameter": "ტროპონინი",
                "reference_range": "უარყოფითი"
            }
        ]
    },
    "footer_note": {
        "equipment": "გამოკვლევა ჩატარდა ანალიზატორ Firance FS-113 _ზე"
    }
}


def set_cell_shading(cell, color):
    shading_elm = OxmlElement('w:shd');
    shading_elm.set(qn('w:fill'), color);
    cell._tc.get_or_add_tcPr().append(shading_elm)


# ========== ROUTES ==========
@app.route('/')
def index(): return render_template('index.html')


@app.route('/search')
def search():
    query = request.args.get('q', '').lower().strip()
    db = load_database()
    results = [p for p in db["patients"] if query in p["last_name"].lower() or query in p["first_name"].lower()]
    return jsonify({"results": sorted(results, key=lambda x: x["created_at"], reverse=True)})


@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(get_saved_docs_folder(), filename)
    return send_file(path, as_attachment=True) if os.path.exists(path) else ("Not Found", 404)


@app.route('/delete/<int:record_id>', methods=['POST'])
def delete_record(record_id):
    db = load_database()
    for i, p in enumerate(db["patients"]):
        if p["id"] == record_id:
            path = os.path.join(get_saved_docs_folder(), p["filename"])
            if os.path.exists(path): os.remove(path)
            db["patients"].pop(i);
            save_database(db);
            return jsonify({"success": True})
    return jsonify({"success": False})


# ========== CBC FUNCTIONS ==========
@app.route('/cbc')
def cbc_form(): return render_template('form_cbc.html', template=CBC_TEMPLATE)


def create_cbc_document(form_data):
    doc = Document()

    # მარჯინები - კომპაქტური
    for section in doc.sections:
        section.top_margin = Cm(0.8)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(1.2)
        section.right_margin = Cm(1.2)

    # ჰედერი - 14pt
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.space_after = Pt(0)
    run = h.add_run("PREMIUM MEDI / პრემიუმ მედი")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 100, 0)

    # სუბტიტრი - 9pt
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.paragraph_format.space_after = Pt(4)
    sr = sub.add_run("საოჯახო მედიცინის ცენტრი | ტელ: 558-27-55-51")
    sr.font.size = Pt(9)

    # სათაური - 11pt
    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_p.paragraph_format.space_after = Pt(8)
    t = t_p.add_run("BL6 - სისხლის საერთო ანალიზი CBC")
    t.font.size = Pt(11)
    t.font.bold = True

    # პაციენტის ინფო - 10pt
    info = doc.add_paragraph()
    info.paragraph_format.space_after = Pt(8)
    info.add_run("პაციენტი: ").bold = True
    info.add_run(
        f"{form_data.get('first_name', '')} {form_data.get('last_name', '')}, {form_data.get('age', '')} წ.   ")
    info.add_run("თარიღი: ").bold = True
    info.add_run(form_data.get('test_date', ''))
    for r in info.runs:
        r.font.size = Pt(10)

    # სისხლის ანალიზი სათაური - 10pt
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(2)
    p1_run = p1.add_run("სისხლის საერთო ანალიზი")
    p1_run.bold = True
    p1_run.font.size = Pt(10)

    # ცხრილი 1 - 10pt
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'

    for i, h_txt in enumerate(['აბრევ.', 'პარამეტრი', 'შედეგი', 'ნორმა', 'ერთ.']):
        cell = table.rows[0].cells[i]
        cell.text = h_txt
        set_cell_shading(cell, 'D9E2F3')
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].paragraph_format.space_after = Pt(1)

    for item in CBC_TEMPLATE["cbc_analysis"]:
        row = table.add_row()
        row.cells[0].text = item['abbr']
        row.cells[1].text = item['parameter']
        row.cells[2].text = form_data.get(f"cbc_{item['abbr']}", '')
        row.cells[3].text = item['reference_range']
        row.cells[4].text = item['unit']
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    r.font.size = Pt(10)

    # ლეიკოციტარული ფორმულა სათაური - 10pt
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.space_before = Pt(8)
    p2_run = p2.add_run("ლეიკოციტარული ფორმულა")
    p2_run.bold = True
    p2_run.font.size = Pt(10)

    # ცხრილი 2 - 10pt
    lt = doc.add_table(rows=1, cols=3)
    lt.style = 'Table Grid'
    for i, h_txt in enumerate(['პარამეტრი', 'შედეგი', 'ნორმა']):
        cell = lt.rows[0].cells[i]
        cell.text = h_txt
        set_cell_shading(cell, 'E2F0D9')
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].paragraph_format.space_after = Pt(1)

    for idx, item in enumerate(CBC_TEMPLATE["leukocyte_formula"]):
        row = lt.add_row()
        row.cells[0].text = item['parameter']
        row.cells[1].text = form_data.get(f'leuko_{idx}', '')
        row.cells[2].text = item['norm']
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    r.font.size = Pt(10)

    # მორფოლოგია - 10pt
    morph = doc.add_paragraph()
    morph.paragraph_format.space_before = Pt(8)
    morph.paragraph_format.space_after = Pt(2)
    morph.add_run("ერითროც. მორფოლოგია: ").bold = True
    morph.add_run(form_data.get('erythrocyte_morphology', '') + "  ")
    morph.add_run("ლეიკოც. მორფოლოგია: ").bold = True
    morph.add_run(form_data.get('leukocyte_morphology', ''))
    for r in morph.runs:
        r.font.size = Pt(10)

    # ფუტერი - 10pt
    foot = doc.add_paragraph()
    foot.paragraph_format.space_before = Pt(12)
    foot.add_run("გამოკვლევა შეასრულა: ").bold = True
    foot.add_run(form_data.get('doctor_name', '') + "    ")
    foot.add_run("ხელმოწერა: _________")
    for r in foot.runs:
        r.font.size = Pt(10)

    return doc


@app.route('/cbc/print', methods=['POST'])
def cbc_print_route():
    fd = request.form.to_dict()
    # 1. შენახვა
    doc = create_cbc_document(fd)
    fname = f"CBC_{fd.get('last_name')}_{datetime.now().strftime('%H%M%S')}.docx"
    fpath = os.path.join(get_saved_docs_folder(), fname)
    doc.save(fpath)
    add_patient_record(fd.get('first_name'), fd.get('last_name'), fd.get('age'), 'CBC', fname, fd.get('test_date'))

    # 2. ბეჭდვის HTML
    html = f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CBC</title>
<style>@page{{size:A4;margin:10mm}}body{{font-family:Arial,sans-serif;padding:10px;font-size:15px}}
h1{{color:green;text-align:center;font-size:18px}}h2{{text-align:center;font-size:16px}}
table{{width:100%;border-collapse:collapse;margin:10px 0}}th,td{{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}}th{{background:#D9E2F3}}.leuko th{{background:#E2F0D9}}</style></head><body>
<h1>PREMIUM MEDI / პრემიუმ მედი</h1><p style="text-align:center">ტელ: 558-27-55-51</p><h2>BL6 - სისხლის საერთო ანალიზი CBC</h2>
<p><b>პაციენტი:</b> {fd.get('first_name')} {fd.get('last_name')}, {fd.get('age')} წ. &nbsp;&nbsp; <b>თარიღი:</b> {fd.get('test_date')}</p>
<table><tr><th>აბრევ.</th><th>პარამეტრი</th><th>შედეგი</th><th>ნორმა</th><th>ერთ.</th></tr>'''
    for item in CBC_TEMPLATE["cbc_analysis"]:
        abbr = item['abbr']
        html += f"<tr><td>{abbr}</td><td>{item['parameter']}</td><td><b>{fd.get(f'cbc_{abbr}', '')}</b></td><td>{item['reference_range']}</td><td>{item['unit']}</td></tr>"
    html += '</table><table class="leuko"><tr><th>პარამეტრი</th><th>შედეგი</th><th>ნორმა</th></tr>'
    for idx, item in enumerate(CBC_TEMPLATE["leukocyte_formula"]):
        html += f"<tr><td>{item['parameter']}</td><td><b>{fd.get(f'leuko_{idx}', '')}</b></td><td>{item['norm']}</td></tr>"
    html += f'''</table><p><b>მორფოლოგია:</b> {fd.get('erythrocyte_morphology', '')} | {fd.get('leukocyte_morphology', '')}</p>
<p><b>შეასრულა:</b> {fd.get('doctor_name', '')} &nbsp;&nbsp; ხელმოწერა: __________</p>
<script>window.onload=function(){{setTimeout(function(){{window.print()}},500)}}</script></body></html>'''
    return Response(html, mimetype='text/html')


# ========== URINE FUNCTIONS ==========
@app.route('/urine')
def urine_form(): return render_template('form_urinalysis.html', template=URINE_TEMPLATE)


def create_urine_document(form_data):
    doc = Document()
    for s in doc.sections: s.top_margin = Cm(0.5); s.bottom_margin = Cm(0.5); s.left_margin = Cm(
        1.0); s.right_margin = Cm(1.0)

    h = doc.add_paragraph();
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    h.paragraph_format.space_after = Pt(0)
    run = h.add_run("PREMIUM MEDI / პრემიუმ მედი");
    run.font.size = Pt(16);
    run.font.bold = True;
    run.font.color.rgb = RGBColor(0, 100, 0)

    sub = doc.add_paragraph();
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    sub.paragraph_format.space_after = Pt(2)
    sub.add_run(
        f"{URINE_TEMPLATE['header']['subtitle']} | ტელ: {', '.join(URINE_TEMPLATE['header']['phones'])}").font.size = Pt(
        12)

    t_p = doc.add_paragraph();
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    t_p.paragraph_format.space_after = Pt(4)
    t = t_p.add_run(f"{URINE_TEMPLATE['test_info']['code']} - {URINE_TEMPLATE['test_info']['name']}");
    t.font.size = Pt(14);
    t.font.bold = True

    info = doc.add_paragraph();
    info.paragraph_format.space_after = Pt(4)
    info.add_run("პაციენტი: ").bold = True;
    info.add_run(
        f"{form_data.get('first_name', '')} {form_data.get('last_name', '')}, {form_data.get('age', '')} წ.   ");
    info.add_run("თარიღი: ").bold = True;
    info.add_run(form_data.get('test_date', ''))
    for r in info.runs: r.font.size = Pt(13)

    doc.add_paragraph().add_run("ფიზიკო-ქიმიური თვისებები").bold = True;
    doc.paragraphs[-1].runs[0].font.size = Pt(13)
    t1 = doc.add_table(rows=1, cols=5);
    t1.style = 'Table Grid'
    for i, h_txt in enumerate(['აბრევ.', 'პარამეტრი', 'შედეგი', 'ნორმა', 'ერთ.']):
        cell = t1.rows[0].cells[i];
        cell.text = h_txt;
        set_cell_shading(cell, 'FFF2CC');
        cell.paragraphs[0].runs[0].font.size = Pt(11);
        cell.paragraphs[0].runs[0].font.bold = True;
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    for idx, item in enumerate(URINE_TEMPLATE["physico_chemical"]):
        row = t1.add_row();
        row.cells[0].text = item['abbr'];
        row.cells[1].text = item['parameter'];
        row.cells[2].text = form_data.get(f'phys_{idx}', '');
        row.cells[3].text = item['norm'];
        row.cells[4].text = item['unit']
        for c in row.cells:
            for p in c.paragraphs: p.paragraph_format.space_after = Pt(0);
            for r in c.paragraphs[0].runs: r.font.size = Pt(11)

    doc.add_paragraph().add_run("მიკროსკოპია").bold = True;
    doc.paragraphs[-1].runs[0].font.size = Pt(13)
    mt = doc.add_table(rows=1, cols=4);
    mt.style = 'Table Grid'
    for i, h_txt in enumerate(['ეპითელიუმი', 'შედეგი', 'ცილინდრები', 'შედეგი']):
        cell = mt.rows[0].cells[i];
        cell.text = h_txt;
        set_cell_shading(cell, 'E2EFDA');
        cell.paragraphs[0].runs[0].font.size = Pt(11);
        cell.paragraphs[0].runs[0].font.bold = True;
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    epi, cyl = URINE_TEMPLATE["microscopy"]["epithelium"], URINE_TEMPLATE["microscopy"]["cylinders"]
    for i in range(max(len(epi), len(cyl))):
        row = mt.add_row()
        if i < len(epi): row.cells[0].text = epi[i]['label']; row.cells[1].text = form_data.get(f"epi_{epi[i]['key']}",
                                                                                                '')
        if i < len(cyl): row.cells[2].text = cyl[i]['label']; row.cells[3].text = form_data.get(f"cyl_{cyl[i]['key']}",
                                                                                                '')
        for c in row.cells:
            for p in c.paragraphs: p.paragraph_format.space_after = Pt(0);
            for r in c.paragraphs[0].runs: r.font.size = Pt(11)

    doc.add_paragraph().add_run("სხვა მონაცემები").bold = True;
    doc.paragraphs[-1].runs[0].font.size = Pt(13)
    ot = doc.add_table(rows=1, cols=4);
    ot.style = 'Table Grid'
    for i, h_txt in enumerate(['პარამეტრი', 'შედეგი', 'პარამეტრი', 'შედეგი']):
        cell = ot.rows[0].cells[i];
        cell.text = h_txt;
        set_cell_shading(cell, 'DDEBF7');
        cell.paragraphs[0].runs[0].font.size = Pt(11);
        cell.paragraphs[0].runs[0].font.bold = True;
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    others = URINE_TEMPLATE["microscopy"]["others"]
    for i in range(0, len(others), 2):
        row = ot.add_row();
        row.cells[0].text = others[i]['parameter'];
        row.cells[1].text = form_data.get(f"other_{others[i]['key']}", '')
        if i + 1 < len(others): row.cells[2].text = others[i + 1]['parameter']; row.cells[3].text = form_data.get(
            f"other_{others[i + 1]['key']}", '')
        for c in row.cells:
            for p in c.paragraphs: p.paragraph_format.space_after = Pt(0);
            for r in c.paragraphs[0].runs: r.font.size = Pt(11)

    foot = doc.add_paragraph();
    foot.paragraph_format.space_before = Pt(6)
    foot.add_run(
        f"აპარატურა: {URINE_TEMPLATE['footer']['equipment']}  შეასრულა: {form_data.get('doctor_name', '')}  ხელმოწერა: _________")
    for r in foot.runs: r.font.size = Pt(12)
    return doc


@app.route('/urine/print', methods=['POST'])
def urine_print_route():
    fd = request.form.to_dict()
    ph = ', '.join(URINE_TEMPLATE['header']['phones'])

    # HTML-ის დასაწყისი
    html = f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Urine</title>
<style>@page{{size:A4;margin:10mm}}body{{font-family:Arial,sans-serif;padding:10px;font-size:13px}}
h1{{color:green;text-align:center;font-size:16px}}h2{{text-align:center;font-size:14px}}
table{{width:100%;border-collapse:collapse;margin:5px 0}}
th,td{{border:1px solid #ddd;padding:4px;text-align:left;font-size:11px}}
th{{background:#FFF2CC}} .micro th{{background:#E2EFDA}} .other th{{background:#DDEBF7}}
h3{{font-size:12px;margin:10px 0 5px 0}}</style></head><body>
<h1>PREMIUM MEDI / პრემიუმ მედი</h1><p style="text-align:center">{URINE_TEMPLATE['header']['subtitle']} | ტელ: {ph}</p>
<h2>{URINE_TEMPLATE['test_info']['code']} - {URINE_TEMPLATE['test_info']['name']}</h2>
<p><b>პაციენტი:</b> {fd.get('first_name')} {fd.get('last_name')}, {fd.get('age')} წ. &nbsp;&nbsp; <b>თარიღი:</b> {fd.get('test_date')}</p>

<h3>ფიზიკო-ქიმიური თვისებები</h3>
<table><tr><th>აბრევ.</th><th>პარამეტრი</th><th>შედეგი</th><th>ნორმა</th><th>ერთ.</th></tr>'''

    # 1. ფიზიკო-ქიმიური
    for idx, item in enumerate(URINE_TEMPLATE["physico_chemical"]):
        html += f"<tr><td>{item['abbr']}</td><td>{item['parameter']}</td><td><b>{fd.get(f'phys_{idx}', '')}</b></td><td>{item['norm']}</td><td>{item['unit']}</td></tr>"
    html += '</table>'

    # 2. მიკროსკოპია
    html += '<h3>მიკროსკოპია</h3><table class="micro"><tr><th>ეპითელიუმი</th><th>შედეგი</th><th>ცილინდრები</th><th>შედეგი</th></tr>'
    epi = URINE_TEMPLATE["microscopy"]["epithelium"]
    cyl = URINE_TEMPLATE["microscopy"]["cylinders"]
    for i in range(max(len(epi), len(cyl))):
        el = epi[i]['label'] if i < len(epi) else ''
        ev = fd.get(f"epi_{epi[i]['key']}", '') if i < len(epi) else ''
        cl = cyl[i]['label'] if i < len(cyl) else ''
        cv = fd.get(f"cyl_{cyl[i]['key']}", '') if i < len(cyl) else ''
        html += f"<tr><td>{el}</td><td><b>{ev}</b></td><td>{cl}</td><td><b>{cv}</b></td></tr>"
    html += '</table>'

    # 3. სხვა მონაცემები (სწორედ ეს ნაწილი გაკლდათ)
    html += '<h3>სხვა მონაცემები</h3><table class="other"><tr><th>პარამეტრი</th><th>შედეგი</th><th>პარამეტრი</th><th>შედეგი</th></tr>'
    others = URINE_TEMPLATE["microscopy"]["others"]
    for i in range(0, len(others), 2):
        p1 = others[i]['parameter']
        v1 = fd.get(f"other_{others[i]['key']}", '')

        p2 = ""
        v2 = ""
        if i + 1 < len(others):
            p2 = others[i + 1]['parameter']
            v2 = fd.get(f"other_{others[i + 1]['key']}", '')

        html += f"<tr><td>{p1}</td><td><b>{v1}</b></td><td>{p2}</td><td><b>{v2}</b></td></tr>"
    html += '</table>'

    # ფუტერი
    html += f'''<br><p><b>აპარატურა:</b> {URINE_TEMPLATE["footer"]["equipment"]} &nbsp;&nbsp; 
    <b>შეასრულა:</b> {fd.get("doctor_name", "")} &nbsp;&nbsp; 
    <b>ხელმოწერა:</b> _____________</p>
    <script>window.onload=function(){{setTimeout(function(){{window.print()}},500)}}</script></body></html>'''

    return Response(html, mimetype='text/html')


# ========== CRP FUNCTIONS ==========
@app.route('/crp')
def crp_form(): return render_template('form_crp.html', template=CRP_TEMPLATE)


def create_crp_document(form_data):
    doc = Document()
    for s in doc.sections: s.top_margin = Cm(1.5); s.bottom_margin = Cm(1.5); s.left_margin = Cm(
        2.0); s.right_margin = Cm(2.0)

    h = doc.add_paragraph();
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    h.paragraph_format.space_after = Pt(6)
    run = h.add_run("PREMIUM MEDI / პრემიუმ მედი");
    run.font.size = Pt(20);
    run.font.bold = True;
    run.font.color.rgb = RGBColor(0, 100, 0)

    sub = doc.add_paragraph();
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    sub.paragraph_format.space_after = Pt(12)
    sub.add_run(
        f"{CRP_TEMPLATE['clinic_info']['description']} | ტელ: {', '.join(CRP_TEMPLATE['clinic_info']['phones'])}").font.size = Pt(
        14)

    t_p = doc.add_paragraph();
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    t_p.paragraph_format.space_after = Pt(16)
    t = t_p.add_run(CRP_TEMPLATE['test_details']['title_ge']);
    t.font.size = Pt(18);
    t.font.bold = True

    info = doc.add_paragraph();
    info.paragraph_format.space_after = Pt(16)
    info.add_run("პაციენტი: ").bold = True;
    info.add_run(
        f"{form_data.get('first_name', '')} {form_data.get('last_name', '')}, {form_data.get('age', '')} წ.          ");
    info.add_run("თარიღი: ").bold = True;
    info.add_run(form_data.get('test_date', ''))
    for r in info.runs: r.font.size = Pt(15)

    table = doc.add_table(rows=1, cols=5);
    table.style = 'Table Grid'
    for i, h_txt in enumerate(['კოდი', 'პარამეტრი', 'შედეგი', 'ნორმა', 'ერთეული']):
        cell = table.rows[0].cells[i];
        cell.text = h_txt;
        set_cell_shading(cell, 'E8DAEF');
        cell.paragraphs[0].runs[0].font.size = Pt(14);
        cell.paragraphs[0].runs[0].font.bold = True
    for item in CRP_TEMPLATE["test_results"]:
        row = table.add_row();
        row.cells[0].text = item['code'];
        row.cells[1].text = item['parameter'];
        row.cells[2].text = form_data.get(f"res_{item['code']}", '');
        row.cells[3].text = item['reference_range'];
        row.cells[4].text = item['unit']
        for c in row.cells:
            for r in c.paragraphs[0].runs: r.font.size = Pt(14)

    foot = doc.add_paragraph();
    foot.paragraph_format.space_before = Pt(24)
    foot.add_run(f"გამოკვლევა შეასრულა: {form_data.get('doctor_name', '')}          ხელმოწერა: _______________")
    for r in foot.runs: r.font.size = Pt(14)
    return doc


@app.route('/crp/print', methods=['POST'])
def crp_print_route():
    fd = request.form.to_dict()
    # 1. შენახვა
    doc = create_crp_document(fd)
    fname = f"CRP_{fd.get('last_name')}_{datetime.now().strftime('%H%M%S')}.docx"
    fpath = os.path.join(get_saved_docs_folder(), fname)
    doc.save(fpath)
    add_patient_record(fd.get('first_name'), fd.get('last_name'), fd.get('age'), 'CRP', fname, fd.get('test_date'))

    # 2. ბეჭდვა
    ph = ', '.join(CRP_TEMPLATE['clinic_info']['phones'])
    html = f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CRP</title>
<style>@page{{size:A4;margin:20mm}}body{{font-family:Arial,sans-serif;padding:20px}}
h1{{color:green;text-align:center;font-size:22px}}h2{{text-align:center;font-size:20px;color:#8e44ad}}
p{{margin:10px 0;font-size:16px}}table{{width:100%;border-collapse:collapse;margin:20px 0}}
th,td{{border:1px solid #ddd;padding:12px;text-align:left;font-size:16px}}th{{background:#E8DAEF}}</style></head><body>
<h1>PREMIUM MEDI / პრემიუმ მედი</h1><p style="text-align:center">{CRP_TEMPLATE['clinic_info']['description']} | ტელ: {ph}</p>
<h2>{CRP_TEMPLATE['test_details']['title_ge']}</h2>
<p><b>პაციენტი:</b> {fd.get('first_name')} {fd.get('last_name')}, {fd.get('age')} წ. &nbsp;&nbsp;&nbsp; <b>თარიღი:</b> {fd.get('test_date')}</p>
<table><tr><th>კოდი</th><th>პარამეტრი</th><th>შედეგი</th><th>ნორმა</th><th>ერთეული</th></tr>'''
    for item in CRP_TEMPLATE["test_results"]:
        code = item['code']
        html += f"<tr><td><b>{code}</b></td><td>{item['parameter']}</td><td><b>{fd.get(f'res_{code}', '')}</b></td><td>{item['reference_range']}</td><td>{item['unit']}</td></tr>"
    html += f'''</table><p><b>შეასრულა:</b> {fd.get('doctor_name', '')} &nbsp;&nbsp;&nbsp; ხელმოწერა: _______________</p>
<script>window.onload=function(){{setTimeout(function(){{window.print()}},500)}}</script></body></html>'''
    return Response(html, mimetype='text/html')


# ========== TROPONIN FUNCTIONS ==========
@app.route('/trop')
def trop_form(): return render_template('form_troponin.html', template=TROPONIN_TEMPLATE)


def create_troponin_document(form_data):
    doc = Document()
    for s in doc.sections: s.top_margin = Cm(1.5); s.bottom_margin = Cm(1.5); s.left_margin = Cm(
        2.0); s.right_margin = Cm(2.0)

    h = doc.add_paragraph();
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    h.paragraph_format.space_after = Pt(6)
    r = h.add_run("PREMIUM MEDI / პრემიუმ მედი");
    r.font.size = Pt(20);
    r.font.bold = True;
    r.font.color.rgb = RGBColor(0, 100, 0)

    sub = doc.add_paragraph();
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    sub.paragraph_format.space_after = Pt(12)
    sub.add_run(
        f"{TROPONIN_TEMPLATE['document_info']['clinic_description']} | {TROPONIN_TEMPLATE['document_info']['contact']}").font.size = Pt(
        14)

    t_p = doc.add_paragraph();
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER;
    t_p.paragraph_format.space_after = Pt(16)
    t = t_p.add_run(TROPONIN_TEMPLATE['test_info']['title']);
    t.font.size = Pt(18);
    t.font.bold = True

    info = doc.add_paragraph();
    info.paragraph_format.space_after = Pt(16)
    info.add_run("პაციენტი: ").bold = True;
    info.add_run(
        f"{form_data.get('first_name', '')} {form_data.get('last_name', '')}, {form_data.get('age', '')} წ.          ");
    info.add_run("თარიღი: ").bold = True;
    info.add_run(form_data.get('test_date', ''))
    for r in info.runs: r.font.size = Pt(15)

    table = doc.add_table(rows=1, cols=4);
    table.style = 'Table Grid'
    for i, h_txt in enumerate(['კოდი', 'პარამეტრი', 'შედეგი', 'ნორმა']):
        cell = table.rows[0].cells[i];
        cell.text = h_txt;
        set_cell_shading(cell, 'FDEBD0');
        cell.paragraphs[0].runs[0].font.size = Pt(14);
        cell.paragraphs[0].runs[0].font.bold = True
    for item in TROPONIN_TEMPLATE["test_info"]["results_table"]:
        row = table.add_row();
        row.cells[0].text = item['code'];
        row.cells[1].text = item['parameter'];
        row.cells[2].text = form_data.get("result_value", '');
        row.cells[3].text = item['reference_range']
        for c in row.cells:
            for r in c.paragraphs[0].runs: r.font.size = Pt(14)

    doc.add_paragraph()
    foot1 = doc.add_paragraph();
    foot1.paragraph_format.space_after = Pt(8)
    foot1.add_run("აპარატურა: ").bold = True;
    foot1.add_run(TROPONIN_TEMPLATE["footer_note"]["equipment"])
    for r in foot1.runs: r.font.size = Pt(14)
    foot2 = doc.add_paragraph();
    foot2.paragraph_format.space_before = Pt(8)
    foot2.add_run("გამოკვლევა შეასრულა: ").bold = True;
    foot2.add_run(form_data.get('doctor_name', '') + "          ");
    foot2.add_run("ხელმოწერა: _________________________")
    for r in foot2.runs: r.font.size = Pt(14)
    return doc


@app.route('/trop/print', methods=['POST'])
def trop_print_route():
    fd = request.form.to_dict()
    # 1. შენახვა
    doc = create_troponin_document(fd)
    fname = f"Trop_{fd.get('last_name')}_{datetime.now().strftime('%H%M%S')}.docx"
    fpath = os.path.join(get_saved_docs_folder(), fname)
    doc.save(fpath)
    add_patient_record(fd.get('first_name'), fd.get('last_name'), fd.get('age'), 'Troponin', fname, fd.get('test_date'))

    # 2. ბეჭდვა
    html = f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Troponin</title>
<style>@page{{size:A4;margin:20mm}}body{{font-family:Arial,sans-serif;padding:20px}}
h1{{color:green;text-align:center;font-size:22px}}h2{{text-align:center;font-size:20px;color:#8e44ad}}
p{{margin:10px 0;font-size:16px}}table{{width:100%;border-collapse:collapse;margin:20px 0}}
th,td{{border:1px solid #ddd;padding:12px;text-align:left;font-size:16px}}th{{background:#FDEBD0}}</style></head><body>
<h1>PREMIUM MEDI / პრემიუმ მედი</h1>
<p style="text-align:center">{TROPONIN_TEMPLATE['document_info']['clinic_description']} | {TROPONIN_TEMPLATE['document_info']['contact']}</p>
<h2>{TROPONIN_TEMPLATE['test_info']['title']}</h2>
<p><b>პაციენტი:</b> {fd.get('first_name', '')} {fd.get('last_name', '')}, {fd.get('age', '')} წ. &nbsp;&nbsp;&nbsp; <b>თარიღი:</b> {fd.get('test_date', '')}</p>
<table><tr><th>კოდი</th><th>პარამეტრი</th><th>შედეგი</th><th>ნორმა</th></tr>'''
    for item in TROPONIN_TEMPLATE["test_info"]["results_table"]:
        html += f"<tr><td>{item['code']}</td><td>{item['parameter']}</td><td><b>{fd.get('result_value', '')}</b></td><td>{item['reference_range']}</td></tr>"
    html += f'''</table><p><b>აპარატურა:</b> {TROPONIN_TEMPLATE["footer_note"]["equipment"]}</p>
<p><b>შეასრულა:</b> {fd.get('doctor_name', '')} &nbsp;&nbsp;&nbsp; <b>ხელმოწერა:</b> _________________________</p>
<script>window.onload=function(){{setTimeout(function(){{window.print()}},500)}}</script></body></html>'''
    return Response(html, mimetype='text/html')


# ========== STARTUP ==========
if __name__ == '__main__':
    threading.Timer(1.5, lambda: webbrowser.open('http://127.0.0.1:5000')).start()
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)