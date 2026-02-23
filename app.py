"""
CRM Backend — Google Sheets Integration
"""
import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
import gspread
from gspread.utils import ValueRenderOption, ValueInputOption
from google.oauth2.service_account import Credentials

app = Flask(__name__)
CORS(app)

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1CJER2qjZhMXx2r0rLBU-lrGW0EsIqRTWM2E11Z7pPl4")
SHEET_NAME = os.environ.get("SHEET_NAME", "PLANILHA CENTRAL")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Colunas na ordem exata da planilha
# A=Gênero, B=Nome, C=Escritório, D=E-mail, E=Cidade, F=Data abordagem
# G=Próximo Follow-up (FÓRMULA — só leitura)
# H=Status, I=Observações
# J=Follow-up 1 (FÓRMULA), K=Follow-up 2 (FÓRMULA), L=Follow-up 3 (FÓRMULA)
COLUMNS = [
    "Gênero", "Nome", "Escritório", "E-mail", "Cidade",
    "Data da abordagem", "Próximo Follow-up", "Status",
    "Observações", "Follow-up 1", "Follow-up 2", "Follow-up 3 (Break-up)"
]

# Índices das colunas editáveis (0-based) — nunca tocar em G(6), J(9), K(10), L(11)
EDITABLE_COLS = [0, 1, 2, 3, 4, 5, 7, 8]  # A, B, C, D, E, F, H, I


import re
from datetime import datetime

def normalize_date(val):
    """Converte dd/mm ou yyyy-mm-dd para dd/mm/yyyy que o Sheets entende."""
    if not val:
        return val
    val = val.strip()
    # yyyy-mm-dd (vem do input date do browser)
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', val)
    if m:
        yyyy, mm, dd = m.groups()
        return f"{dd}/{mm}/{yyyy}"
    # dd/mm sem ano
    m = re.match(r'^(\d{1,2})/(\d{1,2})$', val)
    if m:
        dd, mm = m.groups()
        return f"{dd.zfill(2)}/{mm.zfill(2)}/{datetime.now().year}"
    # dd/mm/yyyy já correto
    if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', val):
        return val
    return val

def get_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise Exception("GOOGLE_CREDENTIALS não configurado")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(SHEET_NAME)
    except Exception:
        sheet = spreadsheet.get_worksheet(0)
    return sheet

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

def sheets_serial_to_date(serial):
    """Converte número serial do Google Sheets para dd/mm/yyyy.
    O Sheets usa epoch 30/12/1899. Serial 1 = 01/01/1900.
    Para converter para data real: subtrair 25569 dias do epoch Unix (01/01/1970).
    """
    try:
        n = float(serial)
        if n < 1:
            return None
        import datetime
        # Ajuste: 25569 = dias entre 30/12/1899 e 01/01/1970
        # Além disso o Sheets tem o bug do 29/02/1900, então para datas > 60: subtrair 1
        delta = datetime.timedelta(days=n - 2)  # -2 corrige epoch + bug 1900
        base = datetime.date(1899, 12, 30)
        d = base + delta
        return d.strftime("%d/%m/%Y")
    except:
        return None

@app.route("/prospectos", methods=["GET"])
def get_prospectos():
    try:
        sheet = get_sheet()
        # FORMATTED_VALUE: fórmulas retornam valor calculado (ex: "05/02")
        # mas datas formatadas como "dd/mm" retornam sem o ano ("26/01")
        all_formatted = sheet.get_all_values(
            value_render_option=ValueRenderOption.formatted
        )
        # UNFORMATTED_VALUE: datas retornam como número serial (ex: 46678)
        # isso nos permite reconstruir a data completa com o ano
        all_unformatted = sheet.get_all_values(
            value_render_option=ValueRenderOption.unformatted
        )

        if not all_formatted:
            return jsonify({"data": [], "total": 0})

        # Normaliza headers: remove \n e espaços extras
        headers = [' '.join(h.replace('\r','').split()) for h in all_formatted[0]]
        # Índice da coluna F (Data da abordagem) = 5 (0-based)
        DATE_COL = 5

        records = []
        for idx, row in enumerate(all_formatted[1:]):
            if not any(cell.strip() for cell in row):
                continue
            while len(row) < len(headers):
                row.append("")
            record = {"_row": idx + 2}
            for i, header in enumerate(headers):
                val = row[i] if i < len(row) else ""
                # Para coluna de data: usar valor unformatted (serial) e converter
                if i == DATE_COL:
                    unf_row = all_unformatted[idx + 1] if idx + 1 < len(all_unformatted) else []
                    unf_val = unf_row[i] if i < len(unf_row) else ""
                    # unf_val é um número serial se for data
                    if unf_val and str(unf_val).replace('.','').replace('-','').isdigit():
                        converted = sheets_serial_to_date(unf_val)
                        val = converted if converted else val
                    # se já vier como dd/mm/yyyy ou dd/mm, mantém
                record[header] = val
            records.append(record)

        return jsonify({"data": records, "total": len(records)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos", methods=["POST"])
def add_prospecto():
    try:
        data = request.get_json()
        sheet = get_sheet()

        # Descobre a última linha que tem e-mail preenchido (coluna D = índice 3)
        all_values = sheet.get_all_values(
            value_render_option=ValueRenderOption.formatted
        )
        last_row_with_email = 1  # começa no header
        for idx, row in enumerate(all_values):
            if idx == 0:
                continue  # pula header
            email_val = row[3] if len(row) > 3 else ""
            if email_val.strip():
                last_row_with_email = idx + 1  # 1-based

        # Insere logo abaixo da última linha com e-mail
        insert_at = last_row_with_email + 1
        row_data = []
        for col in COLUMNS:
            val = data.get(col, "")
            if col == "Data da abordagem" and val:
                val = normalize_date(val)
            row_data.append(val)

        sheet.insert_row(
            row_data,
            index=insert_at,
            value_input_option=ValueInputOption.user_entered
        )

        return jsonify({"success": True, "inserted_at": insert_at})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos/<int:sheet_row>", methods=["PUT"])
def update_prospecto(sheet_row):
    try:
        data = request.get_json()
        sheet = get_sheet()
        for i, col in enumerate(COLUMNS):
            if col in data and i in EDITABLE_COLS:
                val = data[col]
                # Normaliza data de abordagem: "20/02" -> "20/02/2026"
                if col == "Data da abordagem" and val:
                    val = normalize_date(val)
                sheet.update_cell(sheet_row, i + 1, val)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos/<int:sheet_row>", methods=["DELETE"])
def delete_prospecto(sheet_row):
    try:
        sheet = get_sheet()
        sheet.delete_rows(sheet_row)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

# ══ VÍDEOS ENVIADOS ══
VIDEO_SHEET_NAME = os.environ.get("VIDEO_SHEET_NAME", "VÍDEOS ENVIADOS")
# A=Nome, B=E-mail, C=Vídeo enviado no dia, D=Demonstração, E=Status, F-J=Follow-ups
VIDEO_COLUMNS = [
    "Nome", "E-mail", "Vídeo enviado no dia", "Demonstração",
    "Status", "Follow-up 1", "Follow-up 2", "Follow-up 3", "Follow-up 4", "Follow-up 5"
]
VIDEO_EDITABLE_COLS = [0, 1, 2, 3, 4]  # Nome, E-mail, Data, Demo, Status
VIDEO_DATE_COL = 2  # Coluna C

def get_video_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise Exception("GOOGLE_CREDENTIALS não configurado")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet.worksheet(VIDEO_SHEET_NAME)

@app.route("/videos", methods=["GET"])
def get_videos():
    try:
        sheet = get_video_sheet()
        all_formatted = sheet.get_all_values(value_render_option=ValueRenderOption.formatted)
        all_unformatted = sheet.get_all_values(value_render_option=ValueRenderOption.unformatted)
        if not all_formatted:
            return jsonify({"data": [], "total": 0})
        headers = [' '.join(h.replace('\r','').split()).rstrip(':') for h in all_formatted[0]]
        records = []
        for idx, row in enumerate(all_formatted[1:]):
            if not any(cell.strip() for cell in row):
                continue
            while len(row) < len(headers):
                row.append("")
            record = {"_row": idx + 2}
            for i, header in enumerate(headers):
                val = row[i] if i < len(row) else ""
                if i == VIDEO_DATE_COL:
                    unf_row = all_unformatted[idx + 1] if idx + 1 < len(all_unformatted) else []
                    unf_val = unf_row[i] if i < len(unf_row) else ""
                    if unf_val and str(unf_val).replace('.','').replace('-','').isdigit():
                        converted = sheets_serial_to_date(unf_val)
                        val = converted if converted else val
                record[header] = val
            records.append(record)
        return jsonify({"data": records, "total": len(records)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/videos", methods=["POST"])
def add_video():
    try:
        data = request.get_json()
        sheet = get_video_sheet()
        all_values = sheet.get_all_values(value_render_option=ValueRenderOption.formatted)
        last_row = 1
        for idx, row in enumerate(all_values):
            if idx == 0: continue
            if any(cell.strip() for cell in row):
                last_row = idx + 1
        insert_at = last_row + 1
        row_data = []
        for col in VIDEO_COLUMNS:
            val = data.get(col, "")
            if col == "Vídeo enviado no dia" and val:
                val = normalize_date(val)
            row_data.append(val)
        sheet.insert_row(row_data, index=insert_at, value_input_option=ValueInputOption.user_entered)
        return jsonify({"success": True, "inserted_at": insert_at})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/videos/<int:sheet_row>", methods=["PUT"])
def update_video(sheet_row):
    try:
        data = request.get_json()
        sheet = get_video_sheet()
        headers_row = sheet.row_values(1)
        headers = [' '.join(h.replace('\r','').split()) for h in headers_row]
        for field, val in data.items():
            if field == '_row': continue
            if field == "Vídeo enviado no dia" and val:
                val = normalize_date(val)
            try:
                col_idx = headers.index(field)
                sheet.update_cell(sheet_row, col_idx + 1, val)
            except ValueError:
                pass
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/videos/<int:sheet_row>", methods=["DELETE"])
def delete_video(sheet_row):
    try:
        sheet = get_video_sheet()
        sheet.delete_rows(sheet_row)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
