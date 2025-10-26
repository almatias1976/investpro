import os
import asyncio
from datetime import datetime, timezone
from typing import Dict, Any, Set, Optional

from fastapi import FastAPI, WebSocket, WebSocketDisconnect, Request, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import uvicorn

# ─────────────────────────────────────────────────────────────
# .env
# ─────────────────────────────────────────────────────────────
load_dotenv()
INGEST_TOKEN = os.getenv("INGEST_TOKEN", "RTD_123456!")
PORT = int(os.getenv("PORT", "10000"))

# ─────────────────────────────────────────────────────────────
# App
# ─────────────────────────────────────────────────────────────
app = FastAPI(title="RTD Backend", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],         # ajuste depois para seu domínio
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────────────────────
# Estado em memória
# ─────────────────────────────────────────────────────────────
# latest: cache por ticker
latest: Dict[str, Dict[str, Any]] = {}
# subscribers: conexões WS por ticker
subscribers: Dict[str, Set[WebSocket]] = {}

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

# ─────────────────────────────────────────────────────────────
# Schemas
# ─────────────────────────────────────────────────────────────
class IngestPayload(BaseModel):
    ticker: str = Field(..., description="Ex: BBAS3")
    price: float = Field(..., description="Preço atual")
    ts: Optional[str] = Field(None, description="ISO8601; se ausente o servidor preenche")

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
async def broadcast(ticker: str, payload: Dict[str, Any]):
    """Envia a atualização para todos os clientes inscritos no ticker."""
    for ws in subscribers.get(ticker, set()).copy():
        try:
            await ws.send_json(payload)
        except Exception:
            subscribers[ticker].discard(ws)

def register_ws(ticker: str, ws: WebSocket):
    subscribers.setdefault(ticker, set()).add(ws)

def unregister_ws(ws: WebSocket):
    for tset in subscribers.values():
        tset.discard(ws)

# ─────────────────────────────────────────────────────────────
# Endpoints HTTP
# ─────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "time": now_iso()}

@app.get("/tickers")
def list_tickers():
    """Retorna os tickers que já receberam ingest e estão em cache."""
    return sorted(list(latest.keys()))

@app.get("/latest")
def get_latest(ticker: Optional[str] = Query(None, description="Ticker opcional (ex: BBAS3)")):
    """
    Sem parâmetro: retorna todas as últimas cotações do cache.
    Com ?ticker=BBAS3: retorna apenas aquele ticker (404 se não existir).
    """
    if ticker is None:
        return latest
    t = ticker.strip().upper()
    if t not in latest:
        raise HTTPException(status_code=404, detail=f"Ticker {t} sem dados.")
    return latest[t]

@app.post("/ingest")
async def ingest(request: Request, data: IngestPayload):
    """Recebe ticks do bridge (Excel RTD)."""
    token = request.headers.get("x-ingest-token")
    if token != INGEST_TOKEN:
        raise HTTPException(status_code=401, detail="Token inválido")

    t = data.ticker.strip().upper()
    ts = data.ts or now_iso()
    payload = {"ticker": t, "price": data.price, "ts": ts}
    latest[t] = payload

    # empurra para quem está conectado no WS
    asyncio.create_task(broadcast(t, payload))
    return {"ok": True, "ticker": t, "price": data.price, "ts": ts}

# ─────────────────────────────────────────────────────────────
# WebSocket
# ─────────────────────────────────────────────────────────────
@app.websocket("/ws")
async def ws_endpoint(ws: WebSocket):
    """
    Protocolo simples:
      - Cliente conecta e envia uma string com o TICKER (ex: "BBAS3").
      - Servidor registra, envia último valor (se houver) e passa a streamar updates.
      - Mensagens subsequentes do cliente podem ser:
          - "PING" (mantém viva)
          - "SUB:<TICKER>" (assina novo)
          - "UNSUB:<TICKER>" (remove)
    """
    await ws.accept()
    try:
        # primeira msg define a assinatura inicial
        first = await ws.receive_text()
        msg = first.strip()
        if msg.upper().startswith("SUB:"):
            t = msg.split(":", 1)[1].strip().upper()
        else:
            t = msg.strip().upper()

        register_ws(t, ws)
        # manda último valor se existir
        if t in latest:
            await ws.send_json(latest[t])

        # loop
        while True:
            try:
                text = await ws.receive_text()
            except WebSocketDisconnect:
                break

            cmd = text.strip()
            if cmd.upper() == "PING":
                await ws.send_text("PONG")
                continue

            if cmd.upper().startswith("SUB:"):
                nt = cmd.split(":", 1)[1].strip().upper()
                register_ws(nt, ws)
                if nt in latest:
                    await ws.send_json(latest[nt])
                continue

            if cmd.upper().startswith("UNSUB:"):
                ut = cmd.split(":", 1)[1].strip().upper()
                if ut in subscribers:
                    subscribers[ut].discard(ws)
                continue

            # se a pessoa mandou um ticker “cru”, troca a assinatura principal
            if len(cmd) <= 8:  # ticker curto
                unregister_ws(ws)
                register_ws(cmd.upper(), ws)
                if cmd.upper() in latest:
                    await ws.send_json(latest[cmd.upper()])

    finally:
        unregister_ws(ws)

# ─────────────────────────────────────────────────────────────
# Run local
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"✅ API on http://127.0.0.1:{PORT}")
    uvicorn.run("main:app", host="0.0.0.0", port=PORT)
