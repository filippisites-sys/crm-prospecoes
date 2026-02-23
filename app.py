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

@app.route("/prospectos", methods=["GET"])
def get_prospectos():
    try:
        sheet = get_sheet()
        # FORMATTED_VALUE garante que fórmulas retornam o valor calculado
        # ex: coluna G com =SE(...) retorna "06/02/2026" e não a fórmula em si
        all_values = sheet.get_all_values(
            value_render_option=ValueRenderOption.formatted
        )
        if not all_values:
            return jsonify({"data": [], "total": 0})

        headers = all_values[0]
        records = []
        for idx, row in enumerate(all_values[1:]):
            if not any(cell.strip() for cell in row):
                continue
            while len(row) < len(headers):
                row.append("")
            record = {"_row": idx + 2}
            for i, header in enumerate(headers):
                record[header] = row[i] if i < len(row) else ""
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
