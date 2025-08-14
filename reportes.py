import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from jinja2 import Template

# === CONFIGURACION ===
GMAIL_USER = "pruebasunisimple@gmail.com"
GMAIL_PASSWORD = "ukpu xgyo wpuc grvo"

archivo_excel = os.path.join("BDD", "CURSO_PISHING.xlsx")
nombre_hoja = "LANZAMIENTO_CLARO(ABRIL)"
nombre_filtro = "Camila Almendra Soto Rosas"
plantilla_html = os.path.join("Plantillas", "sms_usuario.html")
registro_enviados = "correos_enviados/correos_enviados_individual.csv"

# === Preparacion ===
os.makedirs("correos_enviados", exist_ok=True)

# Leer Excel y filtrar usuario
print("📤 Buscando usuario...")
df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
df_filtrado = df[df["name"] == nombre_filtro]

if df_filtrado.empty:
    print(f"❌ No se encontró el usuario '{nombre_filtro}' en la hoja '{nombre_hoja}'")
else:
    persona = df_filtrado.iloc[0]
    historial = df_filtrado[["campaign", "event"]].to_dict(orient="records")

    # Cargar plantilla
    with open(plantilla_html, "r", encoding="utf-8") as f:
        template = Template(f.read())

    # Renderizar HTML
    html_renderizado = template.render(
        nombre=persona["name"],
        correo=persona["email"],
        departamento=persona["department"],
        historial=historial
    )

    # === Envio de correo ===
    contenido = html_renderizado
    contenido += '<br><img src="cid:claroimg" style="width:100%; max-width:600px; margin-top:20px;">'

    msg = MIMEMultipart("related")
    msg["Subject"] = "Resultado Simulacion de Phishing – ClaroVTR"
    msg["From"] = GMAIL_USER
    msg["To"] = persona["email"]

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(contenido, "html"))
    msg.attach(alt)

    if os.path.exists("claro.png"):
        with open("claro.png", "rb") as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header("Content-ID", "<claroimg>")
            mime_img.add_header("Content-Disposition", "inline", filename="claro.png")
            msg.attach(mime_img)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_PASSWORD)
            server.sendmail(GMAIL_USER, persona["email"], msg.as_string())
            print(f"✅ Correo enviado a {persona['email']}")

            # Registrar envío
            if os.path.exists(registro_enviados):
                df_registro = pd.read_csv(registro_enviados)
            else:
                df_registro = pd.DataFrame(columns=["correo"])

            if persona["email"] not in df_registro["correo"].values:
                df_nuevo = pd.DataFrame({"correo": [persona["email"]]})
                df_nuevo.to_csv(registro_enviados, mode="a", header=not os.path.exists(registro_enviados), index=False)

    except Exception as e:
        print(f"❌ Error enviando correo a {persona['email']}: {e}")
