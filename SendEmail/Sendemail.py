"""
SendEmail.py

Envía el resultado de PriceComparison.py por mail.

Reglas:
- Siempre adjunta el Excel.
- Si hay alertas/errores en el resumen, marca el asunto con [ALERTA].
- Lee credenciales del archivo .env en la raíz del proyecto.

Variables esperadas en .env:
    GMAIL_USER=...@gmail.com         # cuenta usada para autenticar SMTP
    GMAIL_APP_PASSWORD=...           # app password de Gmail (16 chars)
    FROM_NAME=Alquimia Economía y Finanzas
    FROM_EMAIL=contacto@alquimiaconsultora.com
    REPLY_TO=contacto@alquimiaconsultora.com
    ALERT_EMAIL=contacto@alquimiaconsultora.com
"""

import os
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

ENV_PATH = PROJECT_DIR / ".env"

COMPARACION_INPUT_PATH = (
    PROJECT_DIR / "PriceComparison" / "output" / "comparacion_precios.xlsx"
)


# =============================================================================
# CONFIG SMTP
# =============================================================================

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465  # SSL


# =============================================================================
# HELPERS
# =============================================================================

def load_env():
    """
    Carga el .env. En Railway las variables van a estar como env vars,
    en local las cargamos del archivo.
    """
    if ENV_PATH.exists():
        load_dotenv(ENV_PATH)
        print(f"[ENV] Cargado desde: {ENV_PATH}")
    else:
        print(f"[ENV] No existe {ENV_PATH}, usando variables de entorno del sistema.")


def get_required_env(key: str) -> str:
    value = os.getenv(key)
    if not value:
        raise RuntimeError(
            f"Falta variable de entorno requerida: {key}. "
            f"Definila en {ENV_PATH} o en el entorno."
        )
    return value.strip()


def leer_resumen_estado() -> pd.DataFrame:
    """
    Lee la hoja RESUMEN_ESTADO del Excel de comparación.
    Si no se puede leer, devuelve DataFrame vacío.
    """
    try:
        df = pd.read_excel(
            COMPARACION_INPUT_PATH,
            sheet_name="RESUMEN_ESTADO",
            engine="openpyxl",
        )
        return df
    except Exception as e:
        print(f"[WARN] No pude leer RESUMEN_ESTADO: {e}")
        return pd.DataFrame()


def leer_errores() -> pd.DataFrame:
    """
    Lee la hoja ERRORES del Excel de comparación.
    """
    try:
        df = pd.read_excel(
            COMPARACION_INPUT_PATH,
            sheet_name="ERRORES",
            engine="openpyxl",
        )
        return df
    except Exception as e:
        print(f"[WARN] No pude leer ERRORES: {e}")
        return pd.DataFrame()


def hay_alertas(resumen: pd.DataFrame) -> bool:
    """
    Devuelve True si en el resumen aparece algún estado distinto de OK.
    Los estados "malos" son ALERTA, SIN_PRECIO_*, ERROR, NO_ENCONTRADO, etc.
    """
    if resumen.empty or "estado" not in resumen.columns:
        return False

    estados_no_ok = resumen[resumen["estado"] != "OK"]
    cantidad_no_ok = estados_no_ok["cantidad"].sum() if not estados_no_ok.empty else 0

    return cantidad_no_ok > 0


def construir_resumen_html(resumen: pd.DataFrame, errores: pd.DataFrame) -> str:
    """
    Construye el cuerpo HTML del mail con un mini-resumen.
    """
    fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    if resumen.empty:
        tabla_resumen = "<p><i>No se pudo leer el resumen de estado.</i></p>"
    else:
        tabla_resumen = resumen.to_html(index=False, border=0, justify="left")

    # Top 20 errores para no inundar el mail
    if errores.empty:
        bloque_errores = "<p><b>No hay alertas ni errores.</b></p>"
    else:
        cols_mostrar = [
            c for c in [
                "hoja_origen",
                "ticker",
                "precio_alquimia",
                "precio_externo",
                "dif_abs",
                "dif_pct",
                "estado",
            ] if c in errores.columns
        ]
        top = errores[cols_mostrar].head(20)
        bloque_errores = (
            f"<p><b>Top {len(top)} alertas / errores (de {len(errores)} totales):</b></p>"
            + top.to_html(index=False, border=0, justify="left")
        )

    html = f"""
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 13px; color: #222;">
        <h2>Comparación de precios — {fecha_str}</h2>

        <h3>Resumen por estado</h3>
        {tabla_resumen}

        <h3>Detalle de alertas</h3>
        {bloque_errores}

        <hr>
        <p style="font-size: 11px; color: #666;">
          Mail generado automáticamente por el pipeline de Alquimia.
          Excel completo adjunto.
        </p>
      </body>
    </html>
    """

    return html


# =============================================================================
# ENVÍO
# =============================================================================

def construir_mensaje(
    from_name: str,
    from_email: str,
    to_email: str,
    reply_to: str,
    asunto: str,
    cuerpo_html: str,
    adjunto_path: Path,
) -> EmailMessage:
    msg = EmailMessage()

    msg["From"] = formataddr((from_name, from_email))
    msg["To"] = to_email
    msg["Reply-To"] = reply_to
    msg["Subject"] = asunto

    # Versión texto plano simple, por si el cliente no renderiza HTML
    msg.set_content(
        "Este mail contiene la comparación de precios diaria. "
        "Si tu cliente de mail no renderiza HTML, abrí el adjunto Excel."
    )

    msg.add_alternative(cuerpo_html, subtype="html")

    if adjunto_path.exists():
        with open(adjunto_path, "rb") as f:
            data = f.read()

        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=adjunto_path.name,
        )
    else:
        print(f"[WARN] No existe el adjunto: {adjunto_path}. Mando el mail sin adjunto.")

    return msg


def enviar(msg: EmailMessage, gmail_user: str, gmail_app_password: str):
    context = ssl.create_default_context()

    print(f"[SMTP] Conectando a {SMTP_HOST}:{SMTP_PORT}...")
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(gmail_user, gmail_app_password)
        server.send_message(msg)

    print("[SMTP] Mail enviado OK.")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("SEND EMAIL")
    print("========================================")

    load_env()

    gmail_user = get_required_env("GMAIL_USER")
    gmail_app_password = get_required_env("GMAIL_APP_PASSWORD")
    from_name = os.getenv("FROM_NAME", "Alquimia").strip()
    from_email = os.getenv("FROM_EMAIL", gmail_user).strip()
    reply_to = os.getenv("REPLY_TO", from_email).strip()
    alert_email = get_required_env("ALERT_EMAIL")

    if not COMPARACION_INPUT_PATH.exists():
        raise FileNotFoundError(
            f"No encontré {COMPARACION_INPUT_PATH}. "
            "Corré primero PriceComparison.py."
        )

    print(f"Adjunto: {COMPARACION_INPUT_PATH}")
    print(f"De: {from_name} <{from_email}> (auth: {gmail_user})")
    print(f"Para: {alert_email}")

    resumen = leer_resumen_estado()
    errores = leer_errores()

    alerta = hay_alertas(resumen)
    fecha_asunto = datetime.now().strftime("%Y-%m-%d")

    if alerta:
        asunto = f"[ALERTA] Comparación de precios {fecha_asunto}"
    else:
        asunto = f"[OK] Comparación de precios {fecha_asunto}"

    print(f"Asunto: {asunto}")
    print(f"Alertas detectadas: {alerta}")

    cuerpo_html = construir_resumen_html(resumen, errores)

    msg = construir_mensaje(
        from_name=from_name,
        from_email=from_email,
        to_email=alert_email,
        reply_to=reply_to,
        asunto=asunto,
        cuerpo_html=cuerpo_html,
        adjunto_path=COMPARACION_INPUT_PATH,
    )

    enviar(msg, gmail_user, gmail_app_password)

    print("\n========================================")
    print("PROCESO TERMINADO")
    print("========================================")


if __name__ == "__main__":
    main()