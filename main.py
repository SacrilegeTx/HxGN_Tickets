import pandas as pd
from dotenv import dotenv_values
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, InvalidSessionIdException
import time

# Chrome Browser
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


def sleepTime(timeout=3):
    time.sleep(timeout)


def webElementClickable(driver, target, timeout=4):
    try:
        WebDriverWait(driver, timeout, poll_frequency=2).until(target).click()
    except TimeoutException as ex:
        print(f"Exception has been thrown. {str(ex)}")
        driver.quit()
    except InvalidSessionIdException as ex:
        print(f"Exception has been thrown. {str(ex)}")
        driver.quit()


def webElementSendKeys(driver, xpathTarget, text, booleanKeyEnter=True):
    input = driver.find_element(By.XPATH, xpathTarget)
    input.clear()
    input.send_keys(Keys.DELETE)

    if booleanKeyEnter:
        input.send_keys(text + Keys.ENTER)
    else:
        input.send_keys(text)


def createTickets(driver):
    """
    La secuencia para acceder a los tickets es a través del menu SideBar
    1. Clic en Menu Tickets
    2. En el desplegable clic en Mis Tickets
    3. Después clic en Nuevo
    """

    # Paso 1
    menuTickets = EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[2]/div[1]/ul/li[5]/a/span[1]")
    )
    webElementClickable(driver, menuTickets, 10)

    # Paso 2
    myTickets = EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[2]/div[1]/ul/li[5]/ul/li[1]/a")
    )
    webElementClickable(driver, myTickets, 10)

    # Paso 3
    newTicket = EC.presence_of_element_located(
        (By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/div/div/div[1]/div/a")
    )
    webElementClickable(driver, newTicket, 10)

    # Cargamos el excel y lo recorremos
    data = pd.read_excel("Tickets_generados.xlsx", index_col=0)

    for index, row in data.iterrows():
        faena = row["faena"]
        equipo = row["equipo"]
        producto = row["producto"]
        reportante = row["reportado por"]
        descripcion = row["descripcion"]
        detalle = row["detalle"]
        estadoTicket = row["estado"]
        ubicacion = row["ubicacion"]
        tipo = row["tipo"]
        solucion = row["solucion"]
        horaIni = row["hora_ini_t"]
        horaFin = row["hora_fin_t"]

        # Seleccionamos el Sitio o Faena
        site = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/button/span[1]",
            )
        )
        webElementClickable(driver, site, 5)
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div/div/div/input",
            faena,
        )

        # Seleccionamos el Equipo
        equipment = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[3]/div/div/button/span[1]",
            )
        )
        webElementClickable(driver, equipment, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[3]/div/div/div/div/input",
            equipo,
        )

        # Seleccionamos el producto
        product = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[4]/div/div/button",
            )
        )
        webElementClickable(driver, product, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[1]/div[4]/div/div/div/div/input",
            producto,
        )

        # Ingresamos el reportante
        webElementSendKeys(driver, '//*[@id="ReportBy"]', reportante, False)

        # Ingresamos la descripcion
        webElementSendKeys(driver, '//*[@id="Subject"]', descripcion, False)

        # Ingresamos la descripcion
        webElementSendKeys(driver, '//*[@id="Content"]', detalle, False)

        # Seleccionamos el Estado del Ticket
        ticketStatus = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div/button",
            )
        )
        webElementClickable(driver, ticketStatus, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div/div/div/input",
            estadoTicket,
        )

        # Seleccionamos la Ubicacion
        location = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[4]/div/div/button",
            )
        )
        webElementClickable(driver, location, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[4]/div/div/div/div/input",
            ubicacion,
        )

        # Seleccionamos el tipo de atencion
        type = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[5]/div/div/button",
            )
        )
        webElementClickable(driver, type, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[5]/div/div/div/div/input",
            tipo,
        )

        # Seleccionamos la Solucion
        solution = EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[6]/div/div/button",
            )
        )
        webElementClickable(driver, solution, 10)
        sleepTime()
        webElementSendKeys(
            driver,
            "/html/body/div[2]/div[2]/div[2]/div/div/form/div[2]/div/div/div[1]/div/div[2]/div[6]/div/div/div/div/input",
            solucion,
        )

        # Ingresamos la hora de inicio
        webElementSendKeys(driver, '//*[@id="StartAt"]', horaIni)

        # Ingresamos la hora de fin
        webElementSendKeys(driver, '//*[@id="CompletedAt"]', horaFin)

        # Guardar Ticket
        saveTicket = EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/div/form/div[3]/div/button")
        )
        webElementClickable(driver, saveTicket, 10)
        sleepTime()

        # Nuevo ticket
        ticket = EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/div/form/div[3]/div/a[1]")
        )
        webElementClickable(driver, ticket, 10)


def login(driver, username, password):
    """
    Esperamos a que carguen los Inputs que vamos a requerir para ingresar
    nuestras credenciales
    """
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "input"))
    )

    # Username
    webElementSendKeys(driver, '//*[@id="Email"]', username, False)

    # Password
    webElementSendKeys(driver, '//*[@id="Password"]', password)


def loadDriver(url):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()), options=options
    )
    driver.maximize_window()
    driver.get(url)

    return driver


def main():
    # Cargamoas las credenciales y url de nuestro archivo .env
    credentials = dotenv_values(".env")
    username = credentials["username"]
    password = credentials["password"]
    url = credentials["url"]

    driver = loadDriver(url)
    login(driver, username, password)
    # Empieza el registro de tickets
    createTickets(driver)
    # driver.quit()


if __name__ == "__main__":
    main()
