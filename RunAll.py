import subprocess
import sys
from pathlib import Path


# =============================================================================
# PATHS
# =============================================================================

PROJECT_DIR = Path(__file__).resolve().parent

SCRIPTS = [
    PROJECT_DIR / "GetPrices" / "GetPricesAlquimia.py",
    PROJECT_DIR / "GetPrices" / "GetPricesBonistas.py",
    PROJECT_DIR / "GetPrices" / "GetPricesYahoo.py",
    PROJECT_DIR / "PriceComparison" / "PriceComparison.py",
    PROJECT_DIR / "SendEmail" / "SendEmail.py",
]


# =============================================================================
# HELPERS
# =============================================================================

def run_script(script_path: Path):
    if not script_path.exists():
        raise FileNotFoundError(f"No encontré el script: {script_path}")

    print("\n" + "=" * 80)
    print(f"Ejecutando: {script_path.name}")
    print(f"Ruta: {script_path}")
    print("=" * 80)

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=PROJECT_DIR,
        text=True,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"Falló el script {script_path.name} con código {result.returncode}"
        )

    print("\n" + "-" * 80)
    print(f"OK: {script_path.name}")
    print("-" * 80)


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("========================================")
    print("RUN ALL VALIDATION")
    print("========================================")
    print(f"PROJECT_DIR: {PROJECT_DIR}")

    for script in SCRIPTS:
        run_script(script)

    print("\n========================================")
    print("PIPELINE TERMINADO OK")
    print("========================================")

    print("\nArchivos esperados:")
    print(PROJECT_DIR / "GetPrices" / "output" / "tickers_cierre.xlsx")
    print(PROJECT_DIR / "GetPrices" / "output" / "bonistas_precios.xlsx")
    print(PROJECT_DIR / "GetPrices" / "output" / "yahoo_precios.xlsx")
    print(PROJECT_DIR / "PriceComparison" / "output" / "comparacion_precios.xlsx")


if __name__ == "__main__":
    main()