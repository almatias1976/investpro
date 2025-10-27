import os
import time
import json
import requests
from datetime import datetime, timezone
from dotenv import load_dotenv
import win32com.client

# ------------------------------------------------------------
# ⚙️ Configuração
# ------------------------------------------------------------
load_dotenv()

API_BASE = os.getenv("API_BASE", "https://investpro-hbqo.onrender.com")  # URL do Render
INGEST_TOKEN = os.getenv("INGEST_TOKEN", "RTD_123456!")
EXCEL_FILE = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")
TICKER_CELL = "A2"
PRICE_CELL = "B2"
STRIKE_CELL = "C2"
VENC_CELL = "D2"
INTERVAL = 5  # segundos entre ciclos

# ------------------------------------------------------------
# 🧠 Funções auxiliares
# ------------------------------------------------------------
def now_iso():
    return datetime.now(timezone.utc).isoformat()

def abrir_excel():
    """Abre o Excel e ativa RTD."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = True
    excel.AutomationSecurity = 1

    wb = excel.Workbooks.Open(EXCEL_FILE, UpdateLinks=True, ReadOnly=False)
    excel.CalculateFullRebuild()
    time.sleep(2)

    print(f"✅ Excel aberto: {wb.FullName}")
    return excel, wb

def escrever_ticker(wb, ticker):
    """Grava o ticker na célula A2 e força recálculo."""
    ws = wb.Worksheets(SHEET_NAME)
    ws.Range(TICKER_CELL).Value = ticker
    wb.Application.CalculateFullRebuild()
    print(f"✏️  Ticker '{ticker}' gravado em {TICKER_CELL}")

def ler_dados_excel(wb):
    """Lê as colunas B (cotação), C (strike) e D (vencimento)."""
    ws = wb.Worksheets(SHEET_NAME)
    preco = ws.Range(PRICE_CELL).Value
    strike = ws.Range(STRIKE_CELL).Value
    venc = ws.Range(VENC_CELL).Value
    return preco, strike, venc

def enviar_dados(ticker, preco, strike, venc):
    """Envia dados atualizados para a API /update."""
    payload = {
        "ticker": ticker,
        "preco": float(preco or 0),
        "strike": str(strike or "-"),
        "vencimento": str(venc or "-"),
    }
    try:
        r = requests.post(
            f"{API_BASE}/update",
            headers={
                "x-ingest-token": INGEST_TOKEN,
                "Content-Type": "application/json",
            },
            data=json.dumps(payload),
            timeout=5,
        )
        if r.status_code == 200:
            print(f"📡 Dados enviados: {payload}")
        else:
            print(f"[WARN] HTTP {r.status_code}: {r.text}")
    except Exception as e:
        print(f"[ERRO Envio] {e}")

# ------------------------------------------------------------
# 🔁 Loop principal
# ------------------------------------------------------------
if __name__ == "__main__":
    print("🚀 Bridge RTD iniciado")
    print(f"📊 Planilha: {EXCEL_FILE}")
    print(f"🌐 Servidor: {API_BASE}")
    print("------------------------------------------------------------")

    excel, wb = abrir_excel()
    ultimo_ticker = None

    while True:
        try:
            # Busca o último ticker enviado pelo Lovable
            r = requests.get(f"{API_BASE}/latest", timeout=5)
            if r.status_code == 200:
                dados = r.json()
                ticker = dados.get("ticker")

                # Se houver novo ticker, grava no Excel
                if ticker and ticker != ultimo_ticker:
                    escrever_ticker(wb, ticker)
                    ultimo_ticker = ticker
                    time.sleep(3)  # tempo para o RTD atualizar

                # Lê dados atualizados
                preco, strike, venc = ler_dados_excel(wb)
                enviar_dados(ticker, preco, strike, venc)
            else:
                print(f"[WARN] HTTP {r.status_code}: {r.text}")

        except Exception as e:
            print(f"[ERRO Loop] {e}")
            try:
                wb.Application.CalculateFullRebuild()
            except Exception:
                pass
            time.sleep(3)

        time.sleep(INTERVAL)
