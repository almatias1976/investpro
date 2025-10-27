import os
import time
import json
import requests
import pythoncom
import win32com.client
from dotenv import load_dotenv
from datetime import datetime, timezone

# ------------------------------------------------------------
# ⚙️ Configuração
# ------------------------------------------------------------
load_dotenv()

API_BASE = os.getenv("API_BASE", "https://investpro-hbqo.onrender.com")
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
def abrir_excel():
    """Abre ou conecta ao Excel e garante RTD ativo"""
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(EXCEL_FILE)
        print(f"✅ Excel conectado: {EXCEL_FILE}")
        return excel, wb
    except Exception as e:
        print(f"[ERRO Excel] {e}")
        raise

def escrever_ticker(wb, ticker):
    """Escreve o ticker e força o recálculo RTD"""
    ws = wb.Worksheets(SHEET_NAME)
    ws.Range(TICKER_CELL).Value = ticker
    wb.Application.CalculateFullRebuild()
    print(f"✏️  Ticker '{ticker}' gravado em {TICKER_CELL}")

def ler_dados_excel(wb):
    """Lê os valores atualizados da planilha"""
    ws = wb.Worksheets(SHEET_NAME)
    preco = ws.Range(PRICE_CELL).Value or 0
    strike = ws.Range(STRIKE_CELL).Value or "-"
    venc = ws.Range(VENC_CELL).Value or "-"
    return preco, strike, venc

def enviar_dados(ticker, preco, strike, venc):
    """Envia dados atualizados de volta à API Render"""
    payload = {
        "preco": float(preco),
        "strike": str(strike),
        "vencimento": str(venc)
    }

    try:
        r = requests.post(
            f"{API_BASE}/update",
            headers={
                "x-ingest-token": INGEST_TOKEN,
                "Content-Type": "application/json"
            },
            data=json.dumps(payload),
            timeout=10
        )
        if r.status_code == 200:
            print(f"📤 Dados enviados: {payload}")
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
    print(f"🌐 API: {API_BASE}")
    print("------------------------------------------------------------")

    excel, wb = abrir_excel()
    ultimo_ticker = None

    while True:
        try:
            pythoncom.CoInitialize()
            # Busca último ticker da API
            r = requests.get(f"{API_BASE}/latest", timeout=10)
            if r.status_code != 200:
                print(f"[WARN] API {r.status_code}")
                continue

            dados = r.json()
            ticker = dados.get("ticker")

            # Se novo ticker, grava e espera RTD atualizar
            if ticker and ticker != ultimo_ticker:
                escrever_ticker(wb, ticker)
                ultimo_ticker = ticker
                time.sleep(4)

            # Lê e envia dados atualizados
            preco, strike, venc = ler_dados_excel(wb)
            enviar_dados(ticker, preco, strike, venc)

        except Exception as e:
            print(f"[ERRO Loop] {e}")
            time.sleep(5)

        finally:
            pythoncom.CoUninitialize()
            time.sleep(INTERVAL)
