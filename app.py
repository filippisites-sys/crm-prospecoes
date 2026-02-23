"""
CRM Backend — Google Sheets Integration
"""
import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
import gspread
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
COLUMNS = [
    "Gênero", "Nome", "Escritório", "E-mail", "Cidade",
    "Data da abordagem", "Próximo Follow-up", "Status",
    "Observações", "Follow-up 1", "Follow-up 2", "Follow-up 3 (Break-up)"
]

# Índices das colunas editáveis manualmente (0-based)
# A=Gênero, B=Nome, C=Escritório, D=E-mail, E=Cidade, F=Data abordagem, H=Status, I=Observações
EDITABLE_COLS = [0, 1, 2, 3, 4, 5, 7, 8]  # A, B, C, D, E, F, H, I

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
    except:
        sheet = spreadsheet.get_worksheet(0)
    return sheet

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/prospectos", methods=["GET"])
def get_prospectos():
    try:
        sheet = get_sheet()
        all_values = sheet.get_all_values()
        if not all_values:
            return jsonify({"data": [], "total": 0})
        headers = all_values[0]
        records = []
        for idx, row in enumerate(all_values[1:]):
            if not any(cell.strip() for cell in row):
                continue
            # Garante que a row tem tamanho suficiente
            while len(row) < len(headers):
                row.append("")
            record = {"_row": idx + 2}  # linha real no Sheets (1-based + header)
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
        row = [data.get(col, "") for col in COLUMNS]
        sheet.append_row(row, value_input_option="USER_ENTERED")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos/<int:sheet_row>", methods=["PUT"])
def update_prospecto(sheet_row):
    try:
        data = request.get_json()
        sheet = get_sheet()
        for i, col in enumerate(COLUMNS):
            if col in data and i in EDITABLE_COLS:
                sheet.update_cell(sheet_row, i + 1, data[col])
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
