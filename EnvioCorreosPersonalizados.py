import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

# Cargar el archivo .env
load_dotenv()

# Configura tu servidor SMTP con los datos proporcionados
smtp_server = os.getenv('EMAIL_HOST')
smtp_port = int(os.getenv('EMAIL_PORT'))
smtp_user = os.getenv('EMAIL_HOST_USER')
smtp_password = os.getenv('EMAIL_HOST_PASSWORD')

# Carga el archivo Excel
archivo_excel = os.getenv('ARCHIVO_EXCEL')
df = pd.read_excel(archivo_excel)

# Función para enviar correos con archivo adjunto
def enviar_correo(to_email, subject, message, adjunto):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_server
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'html'))

        # Adjuntar archivo
        attachment = open(adjunto, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {adjunto.split("/")[-1]}')
        msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        server.quit()
        print(f'Correo enviado a {to_email}')
    except Exception as e:
        print(f'Error enviando correo a {to_email}: {e}')

# Itera sobre el DataFrame y envía los correos
adjunto = os.getenv('ARCHIVO_ADJUNTO')  

for index, row in df.iterrows():
    nombre = row['NOMBRES']
    apellido = row['APELLIDOS']
    correo = row['CORREO INSTITUCIONAL']
    contrasena = row['CONTRASEÑA PROVISIONAL']
    correo_personal = row['CORREO PERSONAL']
    
    asunto = f"SEMESTRE 2024-I | ACTIVACIÓN DE CORREO INSTITUCIONAL"
    html_message = (f"""\
    <html>
    <head></head>
    <body>
        <h2>BIENVENIDOS AL SEMESTRE 2024-I | Pasos para activar cuenta Unitru</h2>
        <p>Estimado/a {nombre} {apellido},</p>
        <p>Esperamos que este mensaje le encuentre bien. Nos complace informarle que se han generado sus credenciales 
        para el acceso al sistema para el semestre 2024-I. A continuación, encontrará sus detalles de acceso:</p>
        <ul>
            <li><b>Correo Institucional:</b> {correo}</li>
            <li><b>Contraseña Provisional:</b> {contrasena}</li>
        </ul>
        <p>Le recomendamos que cambie su contraseña provisional al ingresar por primera vez para asegurar la seguridad de su cuenta.</p>
        <p>Link del Aula virtual: https://epgvirtual2.unitru.edu.pe/login/index.php</p>
        <p><b>Tiene un plazo de 48 horas para activar su cuenta utilizando estas credenciales.</b> Después de este periodo, la cuenta será desactivada temporalmente hasta que contacte con UTIC.</p>
        <p>Adjunto a este correo encontrará un documento con instrucciones detalladas sobre cómo acceder al sistema y 
        cómo cambiar su contraseña provisional.</p>
        <p>Si tiene alguna pregunta o necesita asistencia adicional, no dude en ponerse en contacto con nuestro equipo de soporte.</p>
        <p>Para mas información y novedades pueden unirse al siguiente grupo de WhatsApp: https://chat.whatsapp.com/IaEOB8NVy6s3NtDEb0EaLB</p>

        <p>Saludos cordiales,<br>
        <b>UTIC_EPG<br>
        <b>Horario de Atencion: </b><br>
        Lunes-Viernes de 8:00am a 2:45pm  y Sabados de 8:00am a 1:00pm</b><br>
        Universidad Nacional de Trujillo</p>
    </body>
    </html>
    """)

    
    enviar_correo(correo_personal, asunto, html_message, adjunto)
