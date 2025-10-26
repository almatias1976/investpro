import os
import time
import json
import requests
from datetime import datetime, timezone
from dotenv import load_dotenv
import win32com.client

# ------------------------------------------------------------
# 🔧 Configuração
# ------------------------------------------------------------
load_dotenv()

API_BASE = os.getenv("API_BASE", "http://localhost:10000")
INGEST_TOKEN = os.getenv("INGEST_TOKEN", "RTD_123456!")
EXCEL_FILE = os.getenv("EXCEL_FILE", r"D:\Python\Sistema\RTD\RTD-python.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "RTD")
TICKER_CELL = os.getenv("TICKER_CELL", "A2")
PRICE_CELL = os.getenv("PRICE_CELL", "B2")
INTERVAL = int(os.getenv("INTERVAL", "5"))

# ------------------------------------------------------------
# 🕒 Funções auxiliares
# ------------------------------------------------------------
def now_iso():
    return datetime.now(timezone.utc).isoformat()

def abrir_excel():
    """Abre o Excel via COM e força a ativação de RTD e recálculo completo."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = True
    excel.AutomationSecurity = 1  # 1 = msoAutomationSecurityLow

    wb = excel.Workbooks.Open(
        EXCEL_FILE,
        UpdateLinks=True,
        ReadOnly=False
    )

    # força recálculo total do RTD
    excel.CalculateFullRebuild()
    time.sleep(3)  # dá tempo para o RTD conectar

    print(f"[SUCESSO] Excel conectado com RTD ativo: {wb.FullName}")
    return excel, wb

def ler_valores_excel(wb):
    """Lê ticker e cotação da planilha."""
    try:
        ws = wb.Worksheets(SHEET_NAME)
        ticker = ws.Range(TICKER_CELL).Value
        preco = ws.Range(PRICE_CELL).Value

        # força recálculo se RTD ainda não respondeu
        if preco in (None, "", 0, -2146826246):
            wb.Application.CalculateFullRebuild()
            time.sleep(2)
            preco = ws.Range(PRICE_CELL).Value

        if not ticker:
            print("[AVISO] Célula A2 (ticker) está vazia.")
            return None, None
        if preco in (None, "", 0, -2146826246):
            print(f"[AVISO] RTD ainda não respondeu para {ticker}. Valor: {preco}")
            return ticker, None

        return str(ticker).strip().upper(), float(preco)
    except Exception as e:
        print(f"[ERRO Leitura Excel] {e}")
        return None, None

def enviar_dados(ticker, preco):
    """Envia dados via POST para a API."""
    payload = {"ticker": ticker, "price": preco, "ts": now_iso()}
    try:
        r = requests.post(
            f"{API_BASE}/ingest",
            headers={
                "x-ingest-token": INGEST_TOKEN,
                "Content-Type": "application/json",
            },
            data=json.dumps(payload),
            timeout=5,
        )
        if r.status_code == 200:
            print(f"[OK] {ticker}: {preco:.2f}")
        else:
            print(f"[WARN] HTTP {r.status_code}: {r.text}")
    except requests.exceptions.ConnectionError:
        print("[ERRO] API não está rodando — aguardando...")
    except Exception as e:
        print(f"[ERRO Envio] {e}")

# ------------------------------------------------------------
# ▶️ Loop principal com reconexão automática
# ------------------------------------------------------------
if __name__ == "__main__":
    print("🚀 Bridge RTD iniciada")
    print(f"📊 Lendo planilha: {EXCEL_FILE}")
    print(f"📡 API destino: {API_BASE}")
    print("------------------------------------------------------------")

    excel, wb = None, None

    while True:
        try:
            if wb is None:
                excel, wb = abrir_excel()

            ticker, preco = ler_valores_excel(wb)
            if ticker and preco is not None:
                enviar_dados(ticker, preco)

        except Exception as e:
            print(f"[AVISO] Excel desconectado, tentando reconectar... ({e})")
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
            wb = None
            time.sleep(3)

        time.sleep(INTERVAL)
