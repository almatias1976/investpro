import os
import time
import threading
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import win32com.client
import pythoncom
from dotenv import load_dotenv

# ------------------------------------------------------------
# 🔧 Configuração
# ------------------------------------------------------------
load_dotenv()

EXCEL_PATH = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")
TICKER_CELL = "A2"
PRICE_CELL = "B2"
STRIKE_CELL = "C2"
VENC_CELL = "D2"

# ------------------------------------------------------------
# ⚙️ Inicialização FastAPI
# ------------------------------------------------------------
app = FastAPI(title="InvestPro RTD API", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------------------------------------
# 📘 Modelo de dados recebido
# ------------------------------------------------------------
class IngestData(BaseModel):
    ticker: str

# ------------------------------------------------------------
# 🧠 Classe que gerencia o Excel via COM com segurança
# ------------------------------------------------------------
class ExcelManager:
    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()

    def connect_excel(self):
        """Conecta ou reconecta ao Excel e retorna workbook e planilha."""
        pythoncom.CoInitialize()
        try:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                # print("🔁 Instância ativa do Excel detectada.")
            except Exception:
                excel = win32com.client.Dispatch("Excel.Application")
                # print("🚀 Nova instância do Excel criada.")

            excel.Visible = True
            excel.DisplayAlerts = False

            # Tenta achar workbook aberto
            wb = None
            for w in excel.Workbooks:
                if self.path.lower() in w.FullName.lower():
                    wb = w
                    break
            if wb is None:
                wb = excel.Workbooks.Open(self.path)
                time.sleep(1)

            ws = wb.Worksheets(SHEET_NAME)
            return excel, wb, ws
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erro ao acessar planilha: {e}")
        finally:
            pythoncom.CoUninitialize()

    def write_ticker_and_read(self, ticker: str):
        """Escreve o ticker e retorna valores RTD (preço, strike, vencimento)."""
        with self.lock:
            pythoncom.CoInitialize()
            try:
                excel, wb, ws = self.connect_excel()

                # Escreve o ticker
                ws.Range(TICKER_CELL).Value = ticker.upper()
                wb.Save()

                # Aguarda RTD atualizar
                time.sleep(2)
                excel.CalculateFullRebuild()
                time.sleep(3)

                price = ws.Range(PRICE_CELL).Value
                strike = ws.Range(STRIKE_CELL).Value
                venc = ws.Range(VENC_CELL).Value

                if not price:
                    raise HTTPException(status_code=404, detail="RTD ainda não respondeu.")

                data = {
                    "ticker": ticker.upper(),
                    "price": price,
                    "strike": strike,
                    "vencimento": venc,
                }
                print(f"✅ Dados retornados: {data}")
                return data

            except Exception as e:
                raise HTTPException(status_code=500, detail=str(e))
            finally:
                pythoncom.CoUninitialize()

excel_manager = ExcelManager(EXCEL_PATH)

# ------------------------------------------------------------
# 🧭 Rotas
# ------------------------------------------------------------
@app.on_event("startup")
def startup_event():
    print("✅ Excel inicializado e mantido aberto.")

@app.post("/ingest")
async def ingest(data: IngestData):
    """Recebe o ticker, escreve no Excel e retorna os valores RTD."""
    try:
        result = excel_manager.write_ticker_and_read(data.ticker)
        return {"status": "ok", "data": result}
    except HTTPException as e:
        raise e
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {"message": "InvestPro RTD API ativa"}
