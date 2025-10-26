from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os
import pythoncom
import win32com.client

app = FastAPI(title="RTD Backend", version="1.1.1")

EXCEL_PATH = r"D:\Python\Sistema\RTD\RTD-python.xlsx"

# Variáveis globais para manter Excel aberto
excel_instance = None
wb_instance = None
ws_instance = None


def init_excel():
    """Abre o Excel uma única vez e mantém ele aberto."""
    global excel_instance, wb_instance, ws_instance
    try:
        pythoncom.CoInitialize()
        excel_instance = win32com.client.Dispatch("Excel.Application")
        excel_instance.Visible = True  # Mantém o Excel aberto
        wb_instance = excel_instance.Workbooks.Open(EXCEL_PATH)
        ws_instance = wb_instance.Sheets(1)
        print("✅ Excel inicializado e mantido aberto.")
    except Exception as e:
        print(f"❌ Falha ao iniciar Excel: {e}")
        raise


@app.on_event("startup")
def startup_event():
    """Executado automaticamente quando o servidor FastAPI inicia."""
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Planilha não encontrada em: {EXCEL_PATH}")
    init_excel()


class IngestPayload(BaseModel):
    ticker: str


@app.post("/ingest")
def ingest(payload: IngestPayload):
    """Escreve o ticker na célula A2 e mantém Excel aberto."""
    global ws_instance
    try:
        ws_instance.Range("A2").Value = payload.ticker
        wb_instance.Save()
        print(f"🟢 Ticker '{payload.ticker}' gravado com sucesso no Excel RTD.")
        return {"status": "ok", "ticker": payload.ticker}
    except Exception as e:
        print(f"❌ Erro ao gravar ticker: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
def health():
    return {"status": "ok", "service": "RTD ativo"}


@app.get("/")
def root():
    return {"message": "API RTD operante. Use POST /ingest para enviar ticker."}
