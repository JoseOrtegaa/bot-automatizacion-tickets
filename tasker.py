from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from openpyxl.styles import PatternFill

NAME_WORKBOOK = './registers/test_tasker.xlsx'

# Configuración global del WebDriver
def setupDriver():
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
                    status = getTicketStatus(driver)
                    
                    # Actualizar columnas en la hoja de Excel
                    editStatusExcel(status, i)
                    print(f'Indice:{i} , Status:{status}, Updated:{datetime.now()}')

    except Exception as e:
        print(f'{datetime.now()} - [TASKER - ERROR] {e}')
    finally:
        print(f'{datetime.now()} - [TASKER - Fin executeReadExcel]')

# Método para obtener el estado del ticket
def getTicketStatus(driver):
    try:
        current_status = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.ID, "current_status_ticket"))
        )
        selector_elements = Select(current_status)
        selected_option = selector_elements.first_selected_option
        return selected_option.text
    except TimeoutException as e:
        print(f'{datetime.now()} - [TASKER - ERROR] {str(e)}')

def editStatusExcel(status, i):
    try:
        if(isinstance(status, str)):
            background = PatternFill(start_color="3ee876", end_color="3ee876", fill_type="solid")
            ws[f'B{i}'].value = status
            ws[f'C{i}'].value = datetime.now()
            ws[f'A{i}'].fill = background
            ws[f'B{i}'].fill = background
            ws[f'C{i}'].fill = background
        else:
             background = PatternFill(start_color="e75151", end_color="e75151", fill_type="solid")
             ws[f'B{i}'].value = f'[ERROR] - Pagina no encontrada.'
             ws[f'A{i}'].fill = background
             ws[f'B{i}'].fill = background
             ws[f'C{i}'].fill = background

    except Exception as e:
        print(f'{datetime.now()} - [TASKER - ERROR] {e}')


# Inicio del programa
print(f'{datetime.now()} - [TASKER - Init]')
wb = load_workbook(filename=NAME_WORKBOOK)
ws = wb.active

# Reutilización del WebDriver
driver = setupDriver()

executeReadExcel(ws, driver)

# Guardar y cerrar
driver.quit()

try:
    wb.save(NAME_WORKBOOK)
    print(f'{datetime.now()} - [TASKER - Fin]')
except Exception as e:
    print(f'{datetime.now()} - [TASKER - ERROR {e}]') 