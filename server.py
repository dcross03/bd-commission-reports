from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime
import base64, io, os, anthropic

app = Flask(__name__)
CORS(app)

TEMPLATE_B64 = os.environ.get('TEMPLATE_B64', '')

def parse_date(s):
    if not s: return None
    for fmt in ('%m/%d/%Y','%m/%d/%y','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S','%Y-%m-%d'):
        try: return datetime.strptime(str(s).strip(), fmt)
        except: pass
    return None

def build_report(email, cc, month_label, orders):
    wb = load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)))
    ws = wb.active
    for row_num in range(11, ws.max_row + 1):
        for col in range(1, 9):
            ws.cell(row_num, col).value = None
    ws['G2'] = month_label
    ws['H4'] = f'=SUM(E11:E{10+len(orders)})'
    ws['H7'] = email
    if cc:
        ws['G8'] = 'CC:'
        ws['G8'].font = Font(name='Calibri', size=11, bold=True)
        ws['G8'].alignment = Alignment(horizontal='center')
        ws['H8'] = cc
        ws['H8'].font = Font(name='Calibri', size=11)
    else:
        ws['G8'] = None
        ws['H8'] = None
    CF = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'
    for i, o in enumerate(orders):
        r = 11 + i
        def sc(col, value, numfmt=None, halign=None):
            cell = ws.cell(r, col)
            cell.value = value
            cell.font = Font(name='Calibri', size=11)
            if halign: cell.alignment = Alignment(horizontal=halign)
            if numfmt: cell.number_format = numfmt
        pd = parse_date(o.get('poDate',''))
        dpd = parse_date(o.get('datePaid',''))
        sc(2, pd or o.get('poDate',''), 'mm-dd-yy', 'left')
        sc(3, str(o.get('po','')), None, 'left')
        sc(4, o.get('cust',''), None, 'center')
        sc(5, float(o.get('total',0)), CF, 'center')
        sc(6, dpd or o.get('datePaid',''), 'mm-dd-yy', 'center')
        sc(7, o.get('rg',''), None, 'center')
        sc(8, o.get('rn',''), None, 'center')
    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

@app.route('/build', methods=['POST'])
def build():
    data = request.json
    results = {}
    for rep in data.get('reports', []):
        results[rep['group']] = build_report(
            rep['email'], rep.get('cc'), rep['monthLabel'], rep['orders']
        )
    return jsonify(results)

@app.route('/draft-emails', methods=['POST'])
def draft_emails():
    data = request.json
    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    if not api_key:
        return jsonify({'error': 'ANTHROPIC_API_KEY not set'}), 500

    client = anthropic.Anthropic(api_key=api_key)
    results = []

    for email in data.get('emails', []):
        to = email['to']
        cc = email.get('cc', '')
        subject = email['subject']
        body = email['body']

        prompt = f"Create a Gmail draft with these exact details:\nTo: {to}"
        if cc:
            prompt += f"\nCC: {cc}"
        prompt += f"\nSubject: {subject}\nBody:\n{body}\n\nReply with only: DONE"

        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=500,
                mcp_servers=[{"type": "url", "url": "https://gmail.mcp.claude.com/mcp", "name": "gmail-mcp"}],
                messages=[{"role": "user", "content": prompt}]
            )
            results.append({'group': email['group'], 'status': 'ok'})
        except Exception as e:
            results.append({'group': email['group'], 'status': 'error', 'message': str(e)})

    return jsonify({'results': results})

@app.route('/')
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
