import requests
import pandas as pd

BASE_URL = "https://api.bcra.gob.ar/estadisticas/v4.0"

VARIABLE_IDS = [44, 45, 135, 136, 137]


def get_last_value(id_variable):
    url = f"{BASE_URL}/monetarias/{id_variable}"

    params = {
        "limit": 1  # trae solo el último dato
    }

    response = requests.get(url, params=params)
    response.raise_for_status()

    data = response.json()

    if data["status"] != 200:
        raise Exception(f"Error API: {data}")

    results = data["results"]

    if not results:
        return None

    detalle = results[0]["detalle"]

    if not detalle:
        return None

    ultimo = detalle[0]

    return {
        "idVariable": id_variable,
        "fecha": ultimo["fecha"],
        "valor": ultimo["valor"]
    }


def main():
    resultados = []

    print("Consultando variables...\n")

    for var_id in VARIABLE_IDS:
        try:
            data = get_last_value(var_id)

            if data:
                print(f"ID {var_id} | Fecha: {data['fecha']} | Valor: {data['valor']}")
                resultados.append(data)
            else:
                print(f"ID {var_id} | Sin datos")

        except Exception as e:
            print(f"Error con ID {var_id}: {e}")

    # Guardar a DataFrame (opcional)
    if resultados:
        df = pd.DataFrame(resultados)
        df.to_csv("variables_seleccionadas.csv", index=False)
        print("\nArchivo guardado: variables_seleccionadas.csv")


if __name__ == "__main__":
    main()