
import os
from time import sleep
import pandas as pd
import socket
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver

def element_presence(by, xpath, time, driver):
    element_present = EC.presence_of_element_located((By.XPATH, xpath))
    WebDriverWait(driver, time).until(element_present)
def is_connected():
    try:
        socket.create_connection(("www.google.com", 80))
        return True
    except BaseException:
        is_connected()

def paste_content(content,driver, el):
    driver.execute_script(
        f'''
const text = `{content}`;
const dataTransfer = new DataTransfer();
dataTransfer.setData('text', text);
const event = new ClipboardEvent('paste', {{
  clipboardData: dataTransfer,
  bubbles: true
}});
arguments[0].dispatchEvent(event)
''',el)

def send_text_message(text, driver, txt_box):
    paste_content(text,driver, txt_box)
    txt_box.send_keys("\n")

def send_image_message(image_name, driver,image_box,root):
    image_box.click()
    upload_button = driver.find_element( By.XPATH,'//input[@type="file"]')
    image_file_path = f"{root}\Data\{image_name}"
    upload_button.send_keys(image_file_path)
    sleep(1)
    send_button = driver.find_element(By.XPATH, '//span[@data-testid="send"]')
    send_button.click()

def main(start, stop, filename, tlf_col, check_col, sheet):
    root =os.path.normpath(os.getcwd() + os.sep + os.pardir )
    data = pd.read_excel(f"{root}\data\{filename}", sheet_name= sheet)
    data[tlf_col] = data[tlf_col].astype(int)
    data[tlf_col] = 51*(10**9) + data[tlf_col]
    if check_col not in data:
        data[check_col] = "No Enviado"

    chrome = webdriver.ChromeOptions()
    chrome.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    driver =  webdriver.Chrome(executable_path="..\chromedriver.exe",  options=chrome)
    sleep(20)
    for i in range(start,stop):
        print(i, data.loc[i,tlf_col])
        if(data.loc[i,check_col] == "Enviado"):
            continue
        try:
            is_connected()
            driver.get("https://web.whatsapp.com/send?phone={}&source=&data=#".format(data.loc[i,tlf_col]))
            WebDriverWait(driver, timeout = 10)        
        except:
            pass
        try:
            element_presence( By.XPATH,'//div[@title="Escribe un mensaje aquí"]', 30, driver)
            txt_box = driver.find_element( By.XPATH, '//div[@title="Escribe un mensaje aquí"]')
            image_box = driver.find_element( By.XPATH,"//span[@data-testid='clip']")
            message = """Estimado Empresario Bodeguero 😃
Te enviamos una guía donde se explica como completar el formulario de registro.
"""
            send_text_message(message, driver, txt_box)
            send_image_message("Pasos de ingreso.png",driver, image_box,root)
            sleep(3)
            data.loc[i,check_col] = "Enviado"
        except Exception as e:
            print("Mensaje no enviado")
            print(e)
            
               
        
    data[tlf_col] = data[tlf_col] - 51*(10**9)
    data.to_excel(f"{root}\data\{filename}", sheet_name = sheet, index = False)
if __name__ == "__main__":

    main(0, 3, "BD Bodega Maestra Base para envios.xlsx", "Teléfono","new_check", "BD envio")
