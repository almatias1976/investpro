from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os
import pythoncom
import win32com.client

app = FastAPI(title="RTD Backend", version="1.1.0")

EXCEL_PATH = r"D:\Python\Sistema\RTD\RTD-python.xlsx"

# Variável global para manter o Excel vivo
excel_instance = None
wb_instance = None
ws_instance = None


def init_excel():
    """Abre o Excel uma única vez e mantém ele aberto."""
    global excel_instance, wb_instance, ws_instance

    try:
        pythoncom.CoInitialize()
        excel_instance = win32com.client.Dispatch("Excel.Application")
        excel_instance.Visible = True  # ⬅️ Mantém aberto visivelmente
        wb_instance = excel_instance.Workbooks.Open(EXCEL_PATH)
        ws_instance = wb_instance.Sheets(1)
        print("✅ Excel inicializado e mantido ativo.")
    except Exception as e:
        print(f"❌ Falha ao iniciar Excel: {e}")
        raise


@app.on_event("startup")
def startup_event():
    """Executado quando o FastAPI inicia"""
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Planilha não encontrada em: {EXCEL_PATH}")
    init_excel()


class IngestPayload(BaseModel):
    ticker: str


@app.post("/ingest")
def ingest(payload: IngestPayload):
    """Escreve o ticker no Excel e mantém aberto (RTD faz o resto)."""
    global ws_instance

    try:
        ws_instance.Range("A2").Value = payload.ticker
        wb_instance.Save()
        print(f"🟢 Ticker '{payload.ticker}' gravado com sucesso.")
        return {"status": "ok", "ticker": payload.ticker}
    except Exception as e:
        print(f"❌ Erro ao gravar ticker: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
def health():
    return {"status": "ok", "service": "RTD ativo"}


@app.get("/")
def root():
    return {"message": "RTD Backend operante. Use POST /ingest para enviar ticker."}
