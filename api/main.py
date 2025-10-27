from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional

app = FastAPI(
    title="InvestPro RTD",
    version="2.0.0",
    description="API intermediária entre Lovable e Excel RTD"
)

# ------------------------------------------------------------
# 🌐 Configuração CORS (permite conexão com Lovable e local)
# ------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # ou especifique o domínio do Lovable se quiser restringir
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------------------------------------
# 🔐 Autenticação básica por token
# ------------------------------------------------------------
INGEST_TOKEN = "RTD_123456!"

# ------------------------------------------------------------
# 🧠 Armazenamento em memória
# ------------------------------------------------------------
data_store = {
    "ticker": None,
    "preco": None,
    "strike": None,
    "vencimento": None
}

# ------------------------------------------------------------
# 📥 Modelo para requisições
# ------------------------------------------------------------
class IngestRequest(BaseModel):
    ticker: str

class UpdateRequest(BaseModel):
    preco: Optional[float] = None
    strike: Optional[str] = None
    vencimento: Optional[str] = None

# ------------------------------------------------------------
# 🚀 Endpoints
# ------------------------------------------------------------
@app.get("/")
def root():
    return {"status": "ok", "message": "InvestPro RTD API ativa"}

@app.post("/ingest")
def ingest(data: IngestRequest, x_ingest_token: str = Header(None)):
    """Recebe o ticker enviado pelo Lovable e guarda para o bridge"""
    if x_ingest_token != INGEST_TOKEN:
        raise HTTPException(status_code=403, detail="Token inválido")

    data_store["ticker"] = data.ticker.upper()
    print(f"✅ Ticker recebido: {data.ticker.upper()}")
    return {"message": f"Ticker {data.ticker.upper()} recebido com sucesso."}

@app.post("/update")
def update(data: UpdateRequest, x_ingest_token: str = Header(None)):
    """Atualiza dados vindos do Excel RTD (via bridge.py local)"""
    if x_ingest_token != INGEST_TOKEN:
        raise HTTPException(status_code=403, detail="Token inválido")

    if data.preco is not None:
        data_store["preco"] = data.preco
    if data.strike is not None:
        data_store["strike"] = data.strike
    if data.vencimento is not None:
        data_store["vencimento"] = data.vencimento

    print(f"📊 Atualizado via RTD: {data_store}")
    return {"message": "Dados atualizados com sucesso.", "data": data_store}

@app.get("/latest")
def latest():
    """Retorna os últimos dados disponíveis"""
    return data_store

# ------------------------------------------------------------
# 🔄 Execução local (modo debug opcional)
# ------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=10000)
