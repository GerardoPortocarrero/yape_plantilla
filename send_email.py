import os
from bs4 import BeautifulSoup
import win32com.client

''' MAIL PROPERTIES
    | (mail.Subject) (mail.ReceivedTime) (mail.SenderName)       |
    | (mail.SenderEmailAddress) (mail.To) (mail.CC)              |
    | (mail.Body) (mail.Attachments.Count) (mail.CreationTime)   |
    | (mail.LastModificationTime) (mail.EntryID)                 |
'''

# Enviar correo atravéz de outlook
# Enviar correo con firma embebida
def main(
        OUTLOOK,
        DOT,
        APPLICATION,
        MAIL_TO,
        MAIL_CC,
        PROCESSED_FILE,
        SUBJECT,
        FIRM_PATH,  # Ruta completa: 'C:/Ruta/firma.png'
        TODAY
):
    # Crear instancia de Outlook
    outlook = win32com.client.Dispatch(OUTLOOK + DOT + APPLICATION)
    mail = outlook.CreateItem(0)  # 0 = MailItem

    # Obtener solo el nombre del archivo de la firma (para el cid)
    cid = os.path.basename(FIRM_PATH)

    # Construir cuerpo HTML correctamente
    body_html = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <p>Estimados,</p>
        <p>Adjunto planificación de yape {TODAY}.</p>
        <p>Saludos,</p>
        <br>
        <img src="cid:{cid}" width="350"><br>
    </body>
    </html>
    """

    # Adjuntar la firma como imagen embebida (cid)
    attachment = mail.Attachments.Add(FIRM_PATH)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
    )

    # Asunto, destinatarios, etc.
    mail.Subject = SUBJECT
    mail.To = MAIL_TO
    mail.CC = MAIL_CC
    mail.HTMLBody = str(BeautifulSoup(body_html, "html.parser"))

    # Agregar archivo adjunto principal
    mail.Attachments.Add(PROCESSED_FILE)

    # Enviar
    mail.Send()

    print("\n" + "="*60)
    print("📄 Reporte enviado")
    print("-" * 60)
    print(f"📝 Reporte : {PROCESSED_FILE}")
    print(f"✉️ Asunto  : {SUBJECT}")
    print("✅ Enviado exitosamente.")
    print("="*60 + "\n")