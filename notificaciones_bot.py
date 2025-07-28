import requests
from datetime import datetime
from dotenv import load_dotenv
import os

# Cargar variables del archivo .env
load_dotenv()

TOKEN = os.getenv("TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

def enviar_recordatorio():
    hora_actual = datetime.now().strftime("%H:%M")
    mensaje = f"⏰ Recordatorio ({hora_actual}): ¡No olvides registrar tu aprendizaje diario!"
    
    try:
        requests.post(
            f"https://api.telegram.org/bot{TOKEN}/sendMessage",
            data={"chat_id": CHAT_ID, "text": mensaje}
        )
        print("Notificación enviada!")
    except Exception as e:
        print(f"Error al enviar: {e}")

if __name__ == "__main__":
    enviar_recordatorio()
