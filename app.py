import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
import pandas as pd
from selenium.webdriver.chrome.options import Options


# Asegúrate de que el path apunte a tu chromedriver.exe
driver_path = r"C:\Users\User\Desktop\Boomit\desarrollos\scraping-reviews-from-googlemaps\chromedriver.exe"

def get_data(driver, dataStructreType):
    """
    Esta función obtiene el texto principal, puntaje y nombre
    """
    print('Obteniendo datos...')
    more_elements = driver.find_elements("class name", "w8nwRe kyuRq")
    for list_more_element in more_elements:
        list_more_element.click()
    
    xpath_base = '//body/div[2]/div[3]/div[8]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[{}]'
    elements = driver.find_element("xpath", xpath_base.format(9 if dataStructreType == 1 else 8))

    childElement = elements.find_element("xpath", ".//div[1]")
    childElementClassName = childElement.get_attribute("class")
    elements = elements.find_elements("xpath", f'//*[@class="{childElementClassName}"]')

    childElementNameClass = childElement.find_element("xpath", './/div[1]/div[1]/div[2]/div[2]/div[1]/button[1]/div[1]').get_attribute("class")
    childElementTextClass = childElement.find_element("xpath", './/div[1]/div[4]/div[2]/div[1]/span[1]').get_attribute("class")
    childElementScoreClass = childElement.find_element("xpath", './/div[1]/div[1]/div[4]/div[1]/span[1]').get_attribute("class")

    lst_data = []
    for data in elements:
        try:
            name = data.find_element("xpath", f'.//*[@class="{childElementNameClass}"]').text
            score = data.find_element("xpath", f'.//*[@class="{childElementScoreClass}"]').get_attribute("aria-label")
            text = data.find_element("xpath", f'.//*[@class="{childElementTextClass}"]').text
            lst_data.append([name + " from GoogleMaps", text, score[0]])
        except:
            pass

    return lst_data

def ifGDRPNotice(driver):
    if "consent.google.com" in driver.current_url:
        driver.execute_script('document.getElementsByTagName("form")[0].submit()')

def ifPageIsFullyLoaded(driver):
    return driver.execute_script("return document.readyState") != "complete"

def counter(driver):
    dataStructreType = 1
    try:
        result = driver.find_element("xpath", '//body/div[2]/div[3]/div[8]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]').find_element("class name", "fontBodySmall").text
    except:
        dataStructreType = 2
        result = driver.find_element("xpath", '//body/div[2]/div[3]/div[8]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]').find_element("class name", "fontBodySmall").text

    result = result.replace(",", "").replace(".", "").split(" ")[0].split("\n")[0]
    return int(int(result) / 10) + 1, dataStructreType

def scrolling(driver, count):
    print("Desplazándose...")
    scrollable_div = driver.find_element("xpath", '//body/div[2]/div[3]/div[8]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[last()]')

    for _ in range(count):
        driver.execute_script(
            """
            var xpathResult = document.evaluate('//body/div[2]/div[3]/div[8]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            var element = xpathResult.singleNodeValue;
            element.scrollTop = element.scrollHeight;
            """,
            scrollable_div
        )
        time.sleep(3)

def write_to_xlsx(data):
    print("Escribiendo a Excel...")
    df = pd.DataFrame(data, columns=["name", "comment", "rating"])
    df.to_excel("out.xlsx", index=False)

if __name__ == "__main__":
    print("Iniciando...")

    # Crear una instancia de las opciones de Chrome
    options = Options()
    options.add_argument("--headless")  # Si quieres ejecutar sin abrir la ventana del navegador (modo "headless")
    options.add_argument("--disable-gpu")  # A veces necesario en modo headless
    options.add_argument("--no-sandbox")  # También útil en entornos de contenedores como Docker

    # Usar el chromedriver con el path correcto
    driver = webdriver.Chrome(service=Service(driver_path), options=options)

    driver.get("https://www.google.com/maps/place/Banco+LaFise/@14.1050648,-87.2066188,17z/data=!4m8!3m7!1s0x8f6fa2be48e4768f:0x75f1d75400270d2d!8m2!3d14.1050596!4d-87.2040439!9m1!1b1!16s%2Fg%2F11c0r4j2jf?entry=ttu&g_ep=EgoyMDI1MDIxMi4wIKXMDSoASAFQAw%3D%3D")  # Reemplaza con la URL deseada

    while ifPageIsFullyLoaded(driver):
        time.sleep(1)

    ifGDRPNotice(driver)

    while ifPageIsFullyLoaded(driver):
        time.sleep(1)

    counter_value = counter(driver)
    scrolling(driver, counter_value[0])

    data = get_data(driver, counter_value[1])
    driver.quit()

    write_to_xlsx(data)
    print("¡Hecho!")
