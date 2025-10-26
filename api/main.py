import os
import time
import threading
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv

# ------------------------------------------------------------
# 🔧 Configuração
# ------------------------------------------------------------
load_dotenv()

EXCEL_FILE = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")

# ------------------------------------------------------------
# 🚀 Inicialização do FastAPI
# ------------------------------------------------------------
app = FastAPI(title="InvestPro RTD API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------------------------------------
# 📦 Modelo de entrada
# ------------------------------------------------------------
class IngestData(BaseModel):
    ticker: str

# ------------------------------------------------------------
# 🧩 Função condicional (apenas ativa no Windows)
# ------------------------------------------------------------
if os.name == "nt":
    import pythoncom
    import win32com.client

    class ExcelController:
        def __init__(self, path: str, sheet_name: str):
            pythoncom.CoInitialize()
            self.path = path
            self.sheet_name = sheet_name
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = True
            self.wb = self.excel.Workbooks.Open(self.path)
            self.ws = self.wb.Worksheets(self.sheet_name)
            print(f"✅ Excel aberto: {self.path}")

        def write_ticker(self, ticker: str):
            """Escreve o ticker na célula A2 e aguarda RTD atualizar."""
            try:
                self.ws.Range("A2").Value = ticker
                self.excel.CalculateFullRebuild()
                print(f"[INFO] Ticker '{ticker}' enviado para Excel.")
                time.sleep(2)
                price = self.ws.Range("B2").Value
                strike = self.ws.Range("C2").Value
                venc = self.ws.Range("D2").Value
                return {
                    "ticker": ticker,
                    "price": price,
                    "strike": strike,
                    "vencimento": venc,
                }
            except Exception as e:
                raise HTTPException(status_code=500, detail=str(e))

else:
    ExcelController = None  # no Linux/Render

excel_ctrl = None

@app.on_event("startup")
async def startup_event():
    """Executa ao iniciar a API."""
    global excel_ctrl
    if os.name == "nt":
        try:
            excel_ctrl = ExcelController(EXCEL_FILE, SHEET_NAME)
            print("✅ Excel inicializado e mantido aberto.")
        except Exception as e:
            print(f"⚠️ Falha ao iniciar Excel: {e}")
    else:
        print("🌐 Executando em ambiente de servidor (Render) — sem Excel local.")

@app.post("/ingest")
def ingest(data: IngestData):
    """Recebe o ticker e grava no Excel (apenas local)."""
    if os.name != "nt":
        # No Render, apenas retorna confirmação
        return {"message": f"Ticker '{data.ticker}' recebido (modo servidor)."}
    try:
        if excel_ctrl is None:
            raise HTTPException(status_code=500, detail="Excel não inicializado.")
        result = excel_ctrl.write_ticker(data.ticker)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
def root():
    return {"status": "API RTD ativa!"}
