import win32com.client

try:
    print("🔍 Tentando conectar ao servidor RTDTrading...")
    rtd = win32com.client.Dispatch("rtdtrading.rtdserver")
    print("✅ Servidor RTDTrading encontrado e registrado no Windows.")
    print("📡 RTD object type:", type(rtd))
except Exception as e:
    print("❌ Servidor RTDTrading NÃO está acessível via COM.")
    print("Erro:", e)
