"""
bcra_twitter.py
===============
Scrapeaa el perfil @BancoCentral_AR en Twitter/X buscando el tweet
de "Principales Variables" y extrae las reservas internacionales
usando OCR sobre la imagen adjunta.

Requisitos:
    pip install selenium pillow pytesseract opencv-python webdriver-manager
    + Tesseract OCR instalado en el sistema

Uso:
    python bcra_twitter.py
    python bcra_twitter.py --headless      # sin ventana visible
    python bcra_twitter.py --output datos  # carpeta de salida
"""

import argparse
import re
import sys
import time
from pathlib import Path
from datetime import datetime

import requests

# ── Selenium ──────────────────────────────────────────────────────────────────
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Instalá las dependencias: pip install selenium webdriver-manager")
    sys.exit(1)

# ── OCR ───────────────────────────────────────────────────────────────────────
try:
    import cv2
    import numpy as np
    from PIL import Image
    import pytesseract
    # Descomentar si estás en Windows:
    # pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("[WARN] pytesseract/opencv no disponibles — solo se guardará la imagen")

# ── Config ────────────────────────────────────────────────────────────────────
TWITTER_URL  = "https://x.com/search?q=bcra+principales+variables&src=typed_query&f=live"
TWITTER_LOGIN_URL = "https://x.com/i/flow/login"
KEYWORD      = "principales variables"
OUTPUT_DIR   = Path("./output/bcra_twitter")
MAX_SCROLLS  = 8       # cuántas veces scrollear hacia abajo buscando el tweet
SCROLL_PAUSE = 2.5     # segundos entre scrolls

# ── Credenciales desde .env ───────────────────────────────────────────────────
def load_credentials():
    env_path = Path(__file__).parent / ".env"
    creds = {}
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                creds[k.strip()] = v.strip()
    return creds.get("TWITTER_USER"), creds.get("TWITTER_PASS")


# ── Driver ────────────────────────────────────────────────────────────────────
def get_driver(headless: bool, chrome_profile: str = None):
    opts = Options()

    if headless:
        opts.add_argument("--headless=new")

    # Usar perfil existente de Chrome (ya logueado en X)
    if chrome_profile:
        opts.add_argument(f"--user-data-dir={chrome_profile}")
        opts.add_argument("--profile-directory=Default")

    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1280,900")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=opts)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


# ── Login ────────────────────────────────────────────────────────────────────
def _click_btn_by_text(driver, *texts, timeout=8):
    """Hace click en el primer botón cuyo span tenga alguno de los textos."""
    xpath = " or ".join([f"text()='{t}'" for t in texts])
    btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, f"//span[{xpath}]/.."))
    )
    btn.click()


def login_twitter(driver, username: str, password: str):
    """Navega a la búsqueda — X redirige al login automáticamente."""
    driver.get(TWITTER_URL)
    time.sleep(6)  # esperar redirección al login

    try:
        # ── PASO 1: campo usuario ─────────────────────────────────────────
        # X usa input[name='text'] para el campo inicial
        user_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='text']"))
        )
        time.sleep(1)
        user_input.click()
        user_input.clear()
        user_input.send_keys(username)
        time.sleep(1)

        # Botón "Siguiente" / "Next"
        _click_btn_by_text(driver, "Siguiente", "Next")
        time.sleep(3)

        # ── PASO 1b: verificación extra (a veces pide username) ───────────
        try:
            verify_input = WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "input[data-testid='ocfEnterTextTextInput']"))
            )
            verify_input.send_keys(username)
            _click_btn_by_text(driver, "Siguiente", "Next")
            time.sleep(3)
        except Exception:
            pass

        # ── PASO 2: contraseña ────────────────────────────────────────────
        pass_input = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
        )
        time.sleep(1)
        pass_input.click()
        pass_input.clear()
        pass_input.send_keys(password)
        time.sleep(1)

        # Botón "Iniciar sesión" / "Log in"
        _click_btn_by_text(driver, "Iniciar sesión", "Log in")
        time.sleep(6)  # esperar redirección post-login
        print("  ✅ Login exitoso")

    except Exception as e:
        print(f"  [WARN] Error en login: {e}")


# ── Buscar tweet ──────────────────────────────────────────────────────────────
def find_bcra_tweet(driver):
    """Navega a la búsqueda y devuelve el primer tweet con imagen."""
    print(f"  Navegando a búsqueda: {TWITTER_URL}...")
    driver.get(TWITTER_URL)
    time.sleep(5)  # esperar que cargue la búsqueda

    seen_ids = set()
    for scroll_i in range(MAX_SCROLLS):
        try:
            articles = driver.find_elements(By.CSS_SELECTOR, "article[data-testid='tweet']")
        except Exception:
            articles = []

        for article in articles:
            try:
                # Usar el texto del tweet para identificarlo
                article_id = article.id
                if article_id in seen_ids:
                    continue
                seen_ids.add(article_id)

                text = article.text.lower()
                # Buscar tweet del BCRA con imagen adjunta
                has_image = bool(article.find_elements(
                    By.CSS_SELECTOR, "img[src*='pbs.twimg.com/media']"))
                is_bcra = "bancoCentral_AR".lower() in text or "banco central" in text

                if has_image and (KEYWORD in text or is_bcra):
                    print(f"  ✅ Tweet encontrado en scroll {scroll_i + 1}")
                    return article
            except Exception:
                continue

        driver.execute_script("window.scrollBy(0, 800);")
        time.sleep(SCROLL_PAUSE)
        print(f"  Scroll {scroll_i + 1}/{MAX_SCROLLS}...", end="\r")

    # Fallback: devolver el primer tweet con imagen que aparezca
    print("\n  ⚠️  No se encontró con keyword — tomando primer tweet con imagen...")
    try:
        articles = driver.find_elements(By.CSS_SELECTOR, "article[data-testid='tweet']")
        for article in articles:
            has_image = bool(article.find_elements(
                By.CSS_SELECTOR, "img[src*='pbs.twimg.com/media']"))
            if has_image:
                print("  ✅ Usando primer tweet con imagen disponible")
                return article
    except Exception:
        pass

    print("  ❌ No se encontró ningún tweet con imagen")
    return None


# ── Extraer imagen ────────────────────────────────────────────────────────────
def get_image_url_from_tweet(article):
    """Extrae la URL de la imagen del tweet."""
    try:
        imgs = article.find_elements(By.CSS_SELECTOR, "img[src*='pbs.twimg.com/media']")
        if imgs:
            url = imgs[0].get_attribute("src")
            # Forzar máxima resolución
            url = re.sub(r'\?.*$', '', url) + "?format=png&name=large"
            return url
    except Exception as e:
        print(f"  [WARN] Error extrayendo imagen: {e}")
    return None


def get_tweet_date(article):
    """Intenta extraer la fecha del tweet."""
    try:
        time_el = article.find_element(By.CSS_SELECTOR, "time")
        return time_el.get_attribute("datetime")[:10]
    except Exception:
        return datetime.now().strftime("%Y-%m-%d")


# ── Descargar imagen ──────────────────────────────────────────────────────────
def download_image(url, path: Path):
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://x.com/"
    }
    try:
        r = requests.get(url, headers=headers, timeout=20)
        r.raise_for_status()
        path.write_bytes(r.content)
        print(f"  Imagen guardada: {path}")
        return True
    except Exception as e:
        print(f"  [ERROR] Descarga fallida: {e}")
        return False


# ── OCR ───────────────────────────────────────────────────────────────────────
def preprocess_image(img_path: Path, out_path: Path):
    """Preprocesa la imagen para mejorar el OCR."""
    img  = cv2.imread(str(img_path))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.convertScaleAbs(gray, alpha=2.0, beta=0)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    cv2.imwrite(str(out_path), thresh)
    return out_path


def extract_text(img_path: Path) -> str:
    img  = Image.open(str(img_path))
    text = pytesseract.image_to_string(img, config="--psm 6")
    return text


def extract_reservas(text: str):
    """
    Busca el número de reservas en el texto OCR.
    El BCRA publica algo como: Reservas internacionales  42.091
    Formato esperado: número con puntos de miles (ej: 43.534)
    """
    lines = text.lower().split("\n")
    for i, line in enumerate(lines):
        if "reserva" in line:
            # Buscar número en la misma línea o la siguiente
            search_text = " ".join(lines[i:i+3])
            numbers = re.findall(r'\b\d{2,3}(?:\.\d{3})+\b', search_text)
            if numbers:
                val = int(numbers[0].replace(".", ""))
                return val, numbers[0]

    # Fallback: primer número grande de 5+ dígitos con puntos
    numbers = re.findall(r'\b\d{2,3}(?:\.\d{3})+\b', text)
    if numbers:
        return int(numbers[0].replace(".", "")), numbers[0]

    return None, None


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--headless",       action="store_true", help="Correr sin ventana")
    ap.add_argument("--output",         default=str(OUTPUT_DIR))
    ap.add_argument("--chrome-profile", default=None,
                    help="Path al perfil de Chrome (ej: C:\\Users\\marco\\AppData\\Local\\Google\\Chrome\\User Data)")
    args = ap.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  BCRA Twitter Scraper")
    print(f"  Buscando: @BancoCentral_AR — Principales Variables")
    print(f"{'='*60}\n")

    driver = get_driver(headless=args.headless, chrome_profile=args.chrome_profile)

    try:
        # 0. Navegar a búsqueda (el perfil ya tiene sesión activa)
        print("[0/4] Navegando a búsqueda de BCRA...")
        driver.get(TWITTER_URL)
        time.sleep(5)

        # Si X redirige al login, intentar loguear
        if "login" in driver.current_url or "i/flow" in driver.current_url:
            print("  Sesión no activa — intentando login...")
            username, password = load_credentials()
            if username and password:
                login_twitter(driver, username, password)
            else:
                print("  ❌ Sin credenciales en .env y sin sesión activa")
                sys.exit(1)

        # 1. Buscar tweet
        print("[1/4] Buscando tweet...")
        article = find_bcra_tweet(driver)
        if not article:
            sys.exit(1)

        tweet_date = get_tweet_date(article)
        print(f"  Fecha del tweet: {tweet_date}")

        # 2. Extraer URL de imagen
        print("\n[2/4] Extrayendo imagen...")
        img_url = get_image_url_from_tweet(article)
        if not img_url:
            print("  ❌ No se encontró imagen en el tweet")
            sys.exit(1)
        print(f"  URL: {img_url}")

        # 3. Descargar imagen
        print("\n[3/4] Descargando imagen...")
        img_path  = output_dir / f"bcra_{tweet_date}.png"
        proc_path = output_dir / f"bcra_{tweet_date}_proc.png"
        if not download_image(img_url, img_path):
            sys.exit(1)

        # 4. OCR
        print("\n[4/4] Aplicando OCR...")
        if not OCR_AVAILABLE:
            print("  [WARN] OCR no disponible — imagen guardada en:", img_path)
            sys.exit(0)

        preprocess_image(img_path, proc_path)
        text = extract_text(proc_path)
        print(f"\n  Texto extraído (primeras 300 chars):\n  {text[:300].strip()}")

        reservas_int, reservas_str = extract_reservas(text)

        print(f"\n{'='*60}")
        if reservas_int:
            print(f"  ✅ Reservas Internacionales: {reservas_str} (USD MM)")
            print(f"  ✅ Valor numérico: {reservas_int:,}")
            print(f"  ✅ Fecha tweet:    {tweet_date}")
        else:
            print("  ❌ No se pudo extraer el valor de reservas del texto OCR")
            print("  Revisá la imagen en:", img_path)
        print(f"{'='*60}\n")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()