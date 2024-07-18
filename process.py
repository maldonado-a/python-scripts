import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

def subir_formulario(driver, info):
    driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "process"))
        )
        driver.find_element(By.NAME, "process").send_keys(info['proceso'])
        driver.find_element(By.NAME, "tipo_riesgo").send_keys(info['tipo_riesgo'])
        driver.find_element(By.NAME, "severidad").send_keys(info['severidad'])
        driver.find_element(By.NAME, "res").send_keys(info['responsable'])
        driver.find_element(By.NAME, "date").send_keys(info['fecha_compromiso'])
        driver.find_element(By.NAME, "obs").send_keys(info['observacion'])

        driver.find_element(By.ID, "submit").click()
        print("Formulario enviado")
    except Exception as e:
        print(f"Error: {e}")

def enviar_email_simulado(responsable, proceso, estado, observacion, fecha_compromiso):
    print(f"\nSimulando envío de correo a {responsable}:")
    print(f"Proceso: {proceso}")
    print(f"Estado: {estado}")
    print(f"Observación: {observacion}")
    print(f"Fecha de Compromiso: {fecha_compromiso}")

def main():
    current_dir = os.path.abspath(os.path.dirname(__file__))
    file_name = "base_seguimiento.xlsx"
    file_path = os.path.join(current_dir, file_name)

    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    try:
        hoja = wb.sheets[0]
        encabezados = hoja.range("A1:J1").value
        encabezados = [encabezado.strip() for encabezado in encabezados]
        print("Encabezados encontrados:", encabezados)

        columnas = {encabezado: idx for idx, encabezado in enumerate(encabezados)}
        print("Columnas mapeadas:", columnas)

        driver = webdriver.Chrome()
        try:
            filas = hoja.range("A2:J" + str(hoja.cells.last_cell.row)).value

            for fila in filas:
                if not fila or fila[columnas["Estado"]] is None:
                    continue
                estado = fila[columnas["Estado"]]

                if estado == "Regularizado":
                    info = {
                        'proceso': fila[columnas["Auditoría/Proceso"]],
                        'tipo_riesgo': fila[columnas["Tipo de Riesgo"]],
                        'severidad': fila[columnas["Severidad\nObservación"]],
                        'responsable': fila[columnas["Responsable"]],
                        'fecha_compromiso': fila[columnas["Fecha\nCompromiso"]].strftime('%Y-%m-%d'),
                        'observacion': fila[columnas["Observación"]],
                    }
                    subir_formulario(driver, info)
                elif estado == "Atrasado":
                    info = {
                        'proceso': fila[columnas["Auditoría/Proceso"]],
                        'estado': estado,
                        'observacion': fila[columnas["Observación"]],
                        'fecha_compromiso': fila[columnas["Fecha\nCompromiso"]].strftime('%Y-%m-%d'),
                        'responsable': fila[columnas["Correo responsable"]],
                    }
                    enviar_email_simulado(info['responsable'], info['proceso'], info['estado'], info['observacion'], info['fecha_compromiso'])
                else:
                    print(f"Estado '{estado}' ignorado para el proceso '{fila[columnas['Auditoría/Proceso']]}'.")
        finally:
            driver.quit()
    finally:
        wb.close()
        app.quit()
        print("Archivo Excel cerrado y aplicación Excel terminada")

if __name__ == "__main__":
    main()
