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
SHEET_NAME = os.environ.get("SHEET_NAME", "Sheet1")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

COLUMNS = [
    "Gênero", "Nome", "Escritório", "E-mail", "Cidade",
    "Data da abordagem", "Próximo Follow-up", "Status",
    "Observações", "Follow-up 1", "Follow-up 2", "Follow-up 3 (Break-up)"
]

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
        records = sheet.get_all_records()
        return jsonify({"data": records, "total": len(records)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos", methods=["POST"])
def add_prospecto():
    try:
        data = request.get_json()
        sheet = get_sheet()
        row = [data.get(col, "") for col in COLUMNS]
        sheet.append_row(row)
        return jsonify({"success": True, "message": "Prospecto adicionado!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos/<int:row_index>", methods=["PUT"])
def update_prospecto(row_index):
    try:
        data = request.get_json()
        sheet = get_sheet()
        # row_index é baseado nos dados (sem header), então +2 para offset do Sheets
        sheet_row = row_index + 2
        for i, col in enumerate(COLUMNS):
            if col in data:
                sheet.update_cell(sheet_row, i + 1, data[col])
        return jsonify({"success": True, "message": "Prospecto atualizado!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/prospectos/<int:row_index>", methods=["DELETE"])
def delete_prospecto(row_index):
    try:
        sheet = get_sheet()
        sheet_row = row_index + 2
        sheet.delete_rows(sheet_row)
        return jsonify({"success": True, "message": "Prospecto removido!"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
