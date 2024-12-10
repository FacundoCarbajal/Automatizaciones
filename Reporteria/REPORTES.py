import glob
import shutil
import threading
import warnings
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import sys
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PyQt5.QtGui import QIcon
from threading import Thread, Event
import re
import win32com.client as client
import pythoncom
import pyautogui
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import pytz
import pandas
import openpyxl
from openpyxl import load_workbook
import os
from pykeepass import  PyKeePass
import win32com.client as win32
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta

def inicializar_driver():
    #user_folder= os.getenv('USERPROFILE')
    profile_path = os.path.join(user_folder,'AppData','Local','Google','Chrome','User Data')  # Ejemplo para Windows
    #profile_path = os.path.join(user_folder,'AppdData','Local','Google','Chrome','User Data')
    chrome_options = Options()
    chrome_options.add_argument(f'user-data-dir={profile_path}')

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except OSError as e:
        print(f"Error al inicializar el driver: {e}")
        sys.exit(1)
    return driver

def Iniciar_Zabbix_Proxys():

    driver.close()

    driver.switch_to.window(driver.window_handles[0])

    try:
        driver.get(Zabbix)

        time.sleep(5)

        login = driver.find_element(By.ID, "login")
        login.click()
        time.sleep(5)

        SAML = driver.find_element(By.CSS_SELECTOR,
                                   "a[href='index_sso.php?request=zabbix.php%3Faction%3Dproxy.list']")
        SAML.click()
        time.sleep(5)
    except:
        print("Zabbix ya esta iniciado")

def Asignacion_de_URLS_Keepass():
    global ZabbixBancor, kp,Diccionario

    URL_ZabbixBancor_KP = kp.find_entries(title='URL_'+Diccionario(1), first=True)


    ZabbixBancor = URL_ZabbixBancor_KP.username if URL_ZabbixBancor_KP else None


    URL_ZabbixBancor_KP = kp.find_entries(title='URL_ZabbixBancor', first=True)


    ZabbixBancor = URL_ZabbixBancor_KP.username if URL_ZabbixBancor_KP else None





    print("Se asignaron URLS_KP")

def obtner_fechas():
    global FechaZabbix1,FechaZabbix2
    ayer=datetime.now()-timedelta(days=1)
    inicio_ayer =ayer.replace(hour=0,minute=0,second=0)
    fin_ayer = ayer.replace(hour=23,minute=59,second=59)
    print("Se cambiaron las fechas")
    return inicio_ayer.strftime("%Y-%m-%d %H:%M:%S"),fin_ayer.strftime("%Y-%m-%d %H:%M:%S")


def Reporte_ZabbixBancor():
    time.sleep(4)
    wait = WebDriverWait(driver, timeout=10)
    Historyboton = wait.until(EC.visibility_of_element_located((By.XPATH,"//label[@for='show_20' and text()='History']")))
    Historyboton.click()
    print("Se paso a Hystori")
    time.sleep(3)

    Datebutton = wait.until(EC.visibility_of_element_located((By.XPATH,"//a[@class='tabfilter-item-link btn-time']")))
    Datebutton.click()
    print("Se paso a las fechas")

    Labelfecha1 = wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='text' and @id='from' and @name='from']")))
    Labelfecha1.clear()
    Labelfecha1.send_keys(FechaZabbix1)
    print("Se agrego info a la primera fecha")

    Labelfecha2 = wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='text' and @id='to' and @name='to']")))
    Labelfecha2.clear()
    Labelfecha2.send_keys(FechaZabbix2)
    print("Se agrego info a la segunda fecha")
    time.sleep(5)

    Applibutton = wait.until(EC.visibility_of_element_located((By.XPATH,"//button[@type='button' and @id='apply' and @name='apply']")))
    Applibutton.click()
    print("Se Aplico")
    time.sleep(3)

    Filtrobutton = wait.until(EC.visibility_of_element_located((By.XPATH,"//a[@aria-label='Home' and contains(@class, 'icon-filter') and contains(@class, 'tabfilter-item-link')]")))
    Filtrobutton.click()
    print("Se paso a Filtro")
    time.sleep(3)

    Applibutton = wait.until(EC.visibility_of_element_located((By.XPATH,"//button[@type='submit' and @name='filter_apply' and @value='1' and text()='Apply']")))
    Applibutton.click()
    print("Se Aplico")
    time.sleep(3)

    DowloadButton = wait.until(EC.visibility_of_element_located((By.XPATH,"//button[@type='button' and @id='export_csv' and @data-url='http://monitoreobancor.bcocba.int/zabbix/zabbix.php?show=2&name=&inventory%5B0%5D%5Bfield%5D=type&inventory%5B0%5D%5Bvalue%5D=&evaltype=0&tags%5B0%5D%5Btag%5D=&tags%5B0%5D%5Boperator%5D=0&tags%5B0%5D%5Bvalue%5D=&show_tags=3&tag_name_format=0&tag_priority=&show_opdata=0&show_timeline=1&filter_name=&filter_show_counter=0&filter_custom_time=0&sort=clock&sortorder=DESC&age_state=0&show_symptoms=0&show_suppressed=0&unacknowledged=0&compact_view=0&details=0&highlight_row=0&action=problem.view.csv']")))
    DowloadButton.click()
    print("Se descargo ")
    time.sleep(5)

    patron_archivo=os.path.join(ruta_descargas,f'zbx_problems_export - {fecha_actual}T*.csv')
    archivo_descargado=None
    while not archivo_descargado:
        archivos = glob.glob(patron_archivo)
        if archivos:
            archivo_descargado=archivos[0]
        else:
            time.sleep(1)
    Nombre_Reporte = FechaZabbix1.replace(':', '').replace(' ', '_') + ".csv"
    ruta_nueva= os.path.join(Reportes,Nombre_Reporte)
    shutil.move(archivo_descargado,ruta_nueva)
    print("Se movio archivo con exito")
    origen = os.path.join(Reportes,Nombre_Reporte)
    destino = os.path.join(user_folder, 'OneDrive - CEDI TECH Consulting', 'Monitoreo 2024', 'Caperta para Script',
                           'Reportes', 'Reportes Bancor', 'Master Reporte', 'Master Reporte.csv')

def copiar_datos_csv(origen, destino):
    global FechaZabbix1, FechaZabbix2, ZabbixBancor, Reportes, Descargas, Nombre_Reporte
    try:
        # Leer el archivo CSV de origen, omitiendo la primera fila
        df_origen = pd.read_csv(origen, skiprows=1)

        # Si el archivo destino ya existe, leerlo para agregar datos
        try:
            df_destino = pd.read_csv(destino)
            # Concatenar los nuevos datos
            df_final = pd.concat([df_destino, df_origen], ignore_index=True)
        except FileNotFoundError:
            # Si el archivo destino no existe, usar solo df_origen
            df_final = df_origen

        # Guardar el DataFrame combinado en el archivo CSV destino
        df_final.to_csv(destino, index=False)
        print("Datos copiados y guardados exitosamente de {} a {}.".format(origen, destino))
    except Exception as e:
        print(f"Error al copiar datos: {e}")


def Correr_thread():
    global FechaZabbix1,FechaZabbix2
    Recorrer_Lista(mi_lista)
    Asignacion_de_URLS_Keepass()
    FechaZabbix1, FechaZabbix2 = obtner_fechas()
    Reporte_ZabbixBancor()
    origen = os.path.join(Reportes, Nombre_Reporte)
    copiar_datos_csv(origen, destino)


def Recorrer_Lista(lista):
    for i in range(len(lista)):  # Iterar por los índices de la lista
        print(f"{lista[i]}")  # Mostrar el elemento actual
        if i + 1 < len(lista):  # Verificar si hay un siguiente elemento
            lista[i] += lista[i + 1]  # Concatenar con el siguiente elemento
        else:
            print("Se terminó la lista")
    print("Lista final:", lista)

if __name__ == "__main__":
    driver = inicializar_driver()
    mi_lista=["Aysam","BF","Citrusvil","Epec","Otek","BLR","Banco Roela"]
    Zabbix=""
    FechaZabbix1=""
    FechaZabbix2 = ""

    #Ruta de DTB Keepass
    user_folder= os.getenv('USERPROFILE')
    ruta_kdbx = os.path.join(user_folder,'OneDrive - CEDI TECH Consulting','Monitoreo 2024','Caperta para Script','DataBaseKeepass','Credenciales.kdbx')
    kp=PyKeePass(ruta_kdbx,password="AsistenteOctubre2024")
    ruta_descargas = os.path.join(user_folder,"Downloads")
    Descargas=ruta_descargas
    ruta_reportes =os.path.join(user_folder,'OneDrive - CEDI TECH Consulting','Monitoreo 2024','Caperta para Script','Reportes','Reportes Bancor')
    Reportes = ruta_reportes
    Nombre_Reporte=""
    destino = os.path.join(user_folder, 'OneDrive - CEDI TECH Consulting', 'Monitoreo 2024', 'Caperta para Script',
                           'Reportes', 'Reportes Bancor', 'Master Reporte', 'Master Reporte.csv')
    fecha_actual=datetime.now().strftime('%Y-%m-%d')
    user_folder = os.getenv('USERPROFILE')
    thread3 = Thread(target=Correr_thread)

    thread3.start()


