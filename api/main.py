import os
import time
import threading
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# Para rodar localmente, este import funciona:
try:
    import win32com.client
except ImportError:
    win32com = None  # Em ambiente Render, win32com não existe

# ------------------------------------------------------------
# 🔧 Configuração
# ------------------------------------------------------------
load_dotenv()

EXCEL_PATH = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")
TICKER_CELL = "A2"
PRICE_CELL = "B2"

app = FastAPI(title="InvestPro RTD", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Libera acesso do Lovable
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

excel = None
wb = None

# ------------------------------------------------------------
# 🧩 Função auxiliar
# ------------------------------------------------------------
def conectar_excel():
    """Conecta ao Excel local ou abre caso necessário."""
    global excel, wb
    if win32com is None:
        print("🌐 Executando em ambiente de servidor (sem Excel).")
        return None, None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(EXCEL_PATH)
        ws = wb.Worksheets(SHEET_NAME)
        print(f"✅ Planilha carregada: {EXCEL_PATH}")
        return ws
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao conectar Excel: {e}")

# ------------------------------------------------------------
# 🚀 Inicialização
# ------------------------------------------------------------
@app.on_event("startup")
async def startup_event():
    global excel, wb
    if win32com:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(EXCEL_PATH)
            print(f"✅ Excel inicializado: {EXCEL_PATH}")
        except Exception as e:
            print(f"⚠️ Não foi possível abrir Excel no startup: {e}")
    else:
        print("🌐 Executando em ambiente de servidor (Render) — sem Excel local.")

# ------------------------------------------------------------
# 📡 Endpoint principal
# ------------------------------------------------------------
@app.post("/ingest")
async def ingest_dados(request: Request):
    """Recebe o ticker e grava na planilha local."""
    data = await request.json()
    ticker = data.get("ticker")

    if not ticker:
        raise HTTPException(status_code=400, detail="Ticker não informado")

    if win32com is None:
        # Ambiente Render — apenas simula resposta
        return {"ticker": ticker, "price": 0.00, "status": "Simulado (Render)"}

    try:
        ws = wb.Worksheets(SHEET_NAME)
        ws.Range(TICKER_CELL).Value = ticker.upper()

        # Aguarda o RTD atualizar o valor da célula B2
        tempo_maximo = 10
        for _ in range(tempo_maximo):
            preco = ws.Range(PRICE_CELL).Value
            if preco not in (None, "", 0, -2146826246):
                break
            time.sleep(1)
            wb.Application.CalculateFullRebuild()

        preco = ws.Range(PRICE_CELL).Value
        if preco in (None, "", 0, -2146826246):
            raise HTTPException(status_code=500, detail="RTD não retornou valor")

        return {"ticker": ticker.upper(), "price": round(float(preco), 2)}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ------------------------------------------------------------
# ✅ Health check
# ------------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "ok"}

