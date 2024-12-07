from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

NAME_WORKBOOK = './registers/test_tasker.xlsx'

# Configuración global del WebDriver
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Ejecutar en modo headless
    options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options=options)
    return driver

# Método para leer Excel y procesar
def executeReadExcel(ws, driver):
    print(f'{datetime.now()} - [TASKER - Init executeReadExcel]')
    try:
        for col in ws.iter_cols(min_row=1, max_col=1):
            for i, cells in enumerate(col, start=1):
                # Verificar si la celda tiene un hipervínculo
                if cells.hyperlink:
                    url_ticket = cells.hyperlink.target
                    driver.get(url_ticket)  # Navegar a la URL
                    status = get_ticket_status(driver)
                    
                    # Actualizar columnas en la hoja de Excel
                    ws[f'B{i}'].value = status
                    ws[f'C{i}'].value = datetime.now()
                    print(f'Indice:{i} , Status:{status}, Updated:{datetime.now()}')

    except Exception as e:
        print(f'{datetime.now()} - [TASKER - ERROR] {e}')
    finally:
        print(f'{datetime.now()} - [TASKER - Fin executeReadExcel]')

# Método para obtener el estado del ticket
def get_ticket_status(driver):
    current_status = WebDriverWait(driver, 1).until(
        EC.presence_of_element_located((By.ID, "current_status_ticket"))
    )
    selector_elements = Select(current_status)
    selected_option = selector_elements.first_selected_option
    return selected_option.text

# Inicio del programa
print(f'{datetime.now()} - [TASKER - Init]')
wb = load_workbook(filename=NAME_WORKBOOK)
ws = wb.active

# Reutilización del WebDriver
driver = setup_driver()

# AQUI COLOCAR UNA EXEPCION
executeReadExcel(ws, driver)

# Guardar y cerrar
driver.quit()
wb.save(NAME_WORKBOOK)
print(f'{datetime.now()} - [TASKER - Fin]')