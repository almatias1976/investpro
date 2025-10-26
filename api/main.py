import os
import time
import threading
import win32com.client
from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# ------------------------------------------------------------
# 🔧 Configuração
# ------------------------------------------------------------
load_dotenv()

EXCEL_FILE = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")
TICKER_CELL = os.getenv("TICKER_CELL", "A2")
PRICE_CELL = os.getenv("PRICE_CELL", "B2")
STRIKE_CELL = os.getenv("STRIKE_CELL", "C2")
VENC_CELL = os.getenv("VENC_CELL", "D2")
INGEST_TOKEN = os.getenv("INGEST_TOKEN", "RTD_123456!")

app = FastAPI(title="RTD Backend", version="2.0.0")

# ------------------------------------------------------------
# 🌍 CORS
# ------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------------------------------------
# 🧠 Controle de instância Excel
# ------------------------------------------------------------
class ExcelManager:
    def __init__(self):
        self.excel = None
        self.wb = None
        self.ws = None
        self._connect()

    def _connect(self):
        """Conecta ou reconecta ao Excel"""
        try:
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = True
            self.wb = self.excel.Workbooks.Open(EXCEL_FILE)
            self.ws = self.wb.Worksheets(SHEET_NAME)
            print(f"✅ Excel inicializado e mantido aberto ({EXCEL_FILE})")
        except Exception as e:
            print(f"❌ Erro ao inicializar Excel: {e}")

    def write_ticker_and_read(self, ticker: str):
        """Escreve o ticker em A2 e retorna valores RTD."""
        try:
            if not self.ws:
                self._connect()

            # escreve o ticker em A2
            self.ws.Range(TICKER_CELL).Value = ticker
            self.wb.Application.CalculateFullRebuild()
            print(f"📩 Ticker '{ticker}' escrito em {TICKER_CELL}")

            # aguarda atualização RTD
            time.sleep(2)

            # lê as células atualizadas
            preco = self.ws.Range(PRICE_CELL).Value
            strike = self.ws.Range(STRIKE_CELL).Value
            venc = self.ws.Range(VENC_CELL).Value

            # retorna dados
            return {
                "ticker": ticker,
                "preco": preco if preco else 0,
                "strike": strike if strike else "--",
                "vencimento": venc if venc else "--",
            }
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erro ao gravar ticker: {e}")

excel_manager = ExcelManager()

# ------------------------------------------------------------
# 🚀 Endpoints
# ------------------------------------------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/ingest")
def ingest(payload: dict, x_ingest_token: str = Header(None)):
    """Recebe o ticker do Lovable, escreve no Excel e retorna valores RTD."""
    if x_ingest_token != INGEST_TOKEN:
        raise HTTPException(status_code=403, detail="Token inválido.")

    ticker = payload.get("ticker")
    if not ticker:
        raise HTTPException(status_code=422, detail="Campo 'ticker' é obrigatório.")

    print(f"🟢 Recebido ticker: {ticker}")

    def background_job():
        try:
            result = excel_manager.write_ticker_and_read(ticker)
            print(f"[OK] RTD retornou: {result}")
        except Exception as e:
            print(f"❌ Erro ao gravar ticker: {e}")

    threading.Thread(target=background_job).start()
    return {"message": f"Ticker {ticker} enviado com sucesso"}

@app.get("/latest")
def latest():
    """Retorna o último ticker e valores RTD"""
    try:
        ws = excel_manager.ws
        ticker = ws.Range(TICKER_CELL).Value
        preco = ws.Range(PRICE_CELL).Value
        strike = ws.Range(STRIKE_CELL).Value
        venc = ws.Range(VENC_CELL).Value
        return {
            "ticker": ticker,
            "preco": preco if preco else 0,
            "strike": strike if strike else "--",
            "vencimento": venc if venc else "--",
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao ler RTD: {e}")
