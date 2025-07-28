import requests
from datetime import datetime

# Configuración de Telegram
TOKEN = "7339579254:AAE3ex1K3oZ5blHRGpnlyt8lz-x1Eh0D0DU"
CHAT_ID = "7646293409"

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