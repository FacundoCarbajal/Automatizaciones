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
from plyer import notification
import pygame
import win32com.client as client
import pythoncom
import pyautogui
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import pytz
import pandas
import openpyxl
import os
from pykeepass import  PyKeePass
import win32com.client as win32
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
import customtkinter as ctk



pygame.mixer.init()
pause_threads = False
Notificaciones_por_alerta = False
stop_threads = False
sonido = r"Elementos\\Notificacion.mp3"

thread3_event = Event()
thread3_event.set()



#-----------------------------------------CREDENCIALES Y VARIABLES----------------------------------

def Asignacion_de_Credenciales():
    global Sigma_User, Sigma_Password, MMG_User, MMG_Pass, Nagios_User, Nagios_Pass

    df = pd.read_excel(credenciales_excel)

    for index, row in df.iterrows():
        try:
            if row['Cliente'] == 'Sigma':
                Sigma_User = str(row['Usuario'])
                Sigma_Password = str(row['Contraseña'])
            elif row['Cliente'] == 'MMG':
                MMG_User = str(row['Usuario'])
                MMG_Pass = str(row['Contraseña'])
            elif row['Cliente'] == 'Nagios':
                Nagios_User = str(row['Usuario'])
                Nagios_Pass = str(row['Contraseña'])
        except Exception as e:
            print(f"No se logró leer la fila {index}: {e}")

    print("Se asignaron credenciales")

def Asignacion_de_Urls():
    global Piru , Zabbix, Zabbix_Aysam,Zabbix_BF,Zabbix_CediCBA,Zabbix_Citrusvil,Zabbix_Epec2,Zabbix_Muni,Zabbix_Otek,Zabbix_Roela,Nagios_Front,Nagios,Sigma,Sigma_Atm,Jira_Backlog,MMG_Login,MMG_Tickets,url_sharepoint

    df = pd.read_excel(Urls)

    for index, row in df.iterrows():
        try:
            if row['Cliente'] == 'Piru':
                Piru = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix':
                Zabbix = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Aysam':
                Zabbix_Aysam = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_BF':
                Zabbix_BF = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_CediCBA':
                Zabbix_CediCBA = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Citrusvil':
                Zabbix_Citrusvil = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Epec2':
                Zabbix_Epec2 = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Mun':
                Zabbix_Muni = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Otek':
                Zabbix_Otek = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Roela':
                Zabbix_Roela = str(row['Urls'])
            elif row['Cliente'] == 'Zabbix_Banco_La_Rioja':
                Zabbix_LR = str(row['Urls'])
            elif row['Cliente'] == 'Nagios_Front':
                Nagios_Front = str(row['Urls'])
            elif row['Cliente'] == 'Nagios':
                Nagios = str(row['Urls'])
            elif row['Cliente'] == 'Sigma':
                Sigma = str(row['Urls'])
            elif row['Cliente'] == 'Sigma_Atm':
                Sigma_Atm = str(row['Urls'])
            elif row['Cliente'] == 'Jira_Backlog':
                Jira_Backlog = str(row['Urls'])
            elif row['Cliente'] == 'MMG_Login':
                MMG_Login = str(row['Urls'])
            elif row['Cliente'] == 'MMG_Tickets':
                MMG_Tickets = str(row['Urls'])
            elif row['Cliente'] == 'url_sharepoint':
                url_sharepoint = str(row['Urls'])
        except Exception as e:
            print(f"No se logró leer la fila {index}: {e}")

    print("Se asignaron Urls")


argentina_tz = pytz.timezone("America/Argentina/Buenos_Aires")

def Asignacion_de_Credenciales_Keepass():
    global Sigma_User, Sigma_Password, MMG_User, MMG_Pass, Nagios_User, Nagios_Pass, kp

    Sigma_KP = kp.find_entries(title='Sigma', first=True)
    Nagios_KP = kp.find_entries(title='Nagios', first=True)
    MMG_KP = kp.find_entries(title='MMG', first=True)

    Sigma_User = Sigma_KP.username if Sigma_KP else None
    Sigma_Password = Sigma_KP.password if Sigma_KP else None

    Nagios_User = Nagios_KP.username if Nagios_KP else None
    Nagios_Pass = Nagios_KP.password if Nagios_KP else None

    MMG_User = MMG_KP.username if MMG_KP else None
    MMG_Pass = MMG_KP.password if MMG_KP else None

    print("Se asignaron Credenciales_KP")

def Asignacion_de_URLS_Keepass():
    global Piru, Zabbix, Zabbix_Aysam, Zabbix_BF, Zabbix_CediCBA, Zabbix_Citrusvil, Zabbix_Epec2, Zabbix_Muni, Zabbix_Otek, Zabbix_Roela,Zabbix_LR, Nagios_Front, Nagios, Sigma, Sigma_Atm, Jira_Backlog, MMG_Login, MMG_Tickets, url_sharepoint, kp

    URL_Piru_KP = kp.find_entries(title='URL_Piru', first=True)
    URL_Zabbix_KP = kp.find_entries(title='URL_Zabbix', first=True)
    URL_Zabbix_Aysam_KP = kp.find_entries(title='URL_Zabbix_Aysam', first=True)
    URL_Zabbix_BF_KP = kp.find_entries(title='URL_Zabbix_BF', first=True)
    URL_Zabbix_CediCBA_KP = kp.find_entries(title='URL_Zabbix_CediCBA', first=True)
    URL_Zabbix_Citrusvil_KP = kp.find_entries(title='URL_Zabbix_Citrusvil', first=True)
    URL_Zabbix_Epec2_KP = kp.find_entries(title='URL_Zabbix_Epec2', first=True)
    URL_Zabbix_Muni_KP = kp.find_entries(title='URL_Zabbix_Muni', first=True)
    URL_Zabbix_Otek_KP = kp.find_entries(title='URL_Zabbix_Otek', first=True)
    URL_Zabbix_Roela_KP = kp.find_entries(title='URL_Zabbix_Roela', first=True)
    URL_Zabbix_BLR_KP = kp.find_entries(title='URL_Zabbix_BLR', first=True)
    URL_Nagios_Front_KP = kp.find_entries(title='URL_Nagios_Front', first=True)
    URL_Nagios_KP = kp.find_entries(title='URL_Nagios', first=True)
    URL_Sigma_KP = kp.find_entries(title='URL_Sigma', first=True)
    URL_Sigma_Atm_KP = kp.find_entries(title='URL_Sigma_Atm', first=True)
    URL_JiraBacklog_KP = kp.find_entries(title='URL_JiraBacklog', first=True)
    URL_MMG_Login_KP = kp.find_entries(title='URL_MMG_Login', first=True)
    URL_MMG_Tickets_KP = kp.find_entries(title='URL_MMG_Tickets', first=True)
    URL_Sharepoint_KP = kp.find_entries(title='URL_Sharepoint', first=True)

    Piru = URL_Piru_KP.username if URL_Piru_KP else None

    Zabbix = URL_Zabbix_KP.username if URL_Zabbix_KP else None

    Zabbix_Aysam = URL_Zabbix_Aysam_KP.username if URL_Zabbix_Aysam_KP else None

    Zabbix_BF = URL_Zabbix_BF_KP.username if URL_Zabbix_BF_KP else None

    Zabbix_CediCBA = URL_Zabbix_CediCBA_KP.username if URL_Zabbix_CediCBA_KP else None

    Zabbix_Citrusvil = URL_Zabbix_Citrusvil_KP.username if URL_Zabbix_Citrusvil_KP else None

    Zabbix_Epec2 = URL_Zabbix_Epec2_KP.username if URL_Zabbix_Epec2_KP else None

    Zabbix_Muni = URL_Zabbix_Muni_KP.username if URL_Zabbix_Muni_KP else None

    Zabbix_Otek = URL_Zabbix_Otek_KP.username if URL_Zabbix_Otek_KP else None

    Zabbix_Roela = URL_Zabbix_Roela_KP.username if URL_Zabbix_Roela_KP else None

    Zabbix_LR = URL_Zabbix_BLR_KP.username if URL_Zabbix_BLR_KP else None

    Nagios_Front = URL_Nagios_Front_KP.username if URL_Nagios_Front_KP else None

    Nagios = URL_Nagios_KP.username if URL_Nagios_KP else None

    Sigma = URL_Sigma_KP.username if URL_Sigma_KP else None

    Sigma_Atm = URL_Sigma_Atm_KP.username if URL_Sigma_Atm_KP else None

    Jira_Backlog = URL_JiraBacklog_KP.username if URL_JiraBacklog_KP else None

    MMG_Login = URL_MMG_Login_KP.username if URL_MMG_Login_KP else None

    MMG_Tickets = URL_MMG_Tickets_KP.username if URL_MMG_Tickets_KP else None

    url_sharepoint = URL_Sharepoint_KP.username if URL_Sharepoint_KP else None




    print("Se asignaron URLS_KP")

#--------------------------------------------BACK----------------------------------
#Funciones extra#
def inicializar_driver():
    user_folder = os.getenv('USERPROFILE')
    # Corrección de la ruta 'AppData' y configuración de user-data-dir
    profile_path = os.path.join(user_folder, 'AppData', 'Local', 'Google', 'Chrome', 'User Data')
    chrome_options = Options()
    chrome_options.add_argument(f'--user-data-dir={profile_path}')  # Asegúrate de agregar el prefijo '--'

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except OSError as e:
        print(f"Error al inicializar el driver: {e}")
        sys.exit(1)

    return driver

def iniciar_tkinter():
    global root, cuadrado1, cuadrado2
    root = ctk.CTk()
    root.geometry("300x200")
    root.title("Monitoreo de Tickets")

    root.attributes("-topmost",True)
    root.resizable(False,False)
    root.overrideredirect(False)

    # Cuadros para mostrar los tickets sin asignar
    cuadrado1 = ctk.CTkLabel(root, text="Tickets sin Asignar Jira: 0", font=("Arial", 14))
    cuadrado1.pack(pady=10)

    cuadrado2 = ctk.CTkLabel(root, text="Tickets sin Asignar MMG: 0", font=("Arial", 14))
    cuadrado2.pack(pady=10)

    root.mainloop()

# Función para actualizar el número de tickets en la interfaz
def actualizar_tickets():
    global cuadrado1, cuadrado2, numero_de_tickets,numero_de_tickets_mmg
    cuadrado1.configure(text=f"Tickets sin Asignar Jira: {numero_de_tickets}")
    cuadrado2.configure(text=f"Tickets sin Asignar MMG: {numero_de_tickets_mmg}")
    numero_de_tickets_mmg=0

def enviar_correo_outlook(destinatario, asunto, cuerpo):
    import pythoncom
    from win32com import client
    pythoncom.CoInitialize()
    try:
        outlook = client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = asunto
        mail.Body = cuerpo
        mail.Send()
    except Exception as e:
        print(f"Error al enviar correo: {e}")
    finally:
        pythoncom.CoUninitialize()

def enviar_correo_en_hilo(destinatario, asunto, cuerpo):
    try:
        correo_thread = Thread(target=enviar_correo_outlook, args=(destinatario, asunto, cuerpo))
        correo_thread.start()
    except Exception as e:
        print(f"Error al iniciar el hilo de correo: {e}")

def create_tray_icon():
    app = QApplication(sys.argv)
    tray_icon = QSystemTrayIcon(QIcon("Elementos\\icono.svg"), app)
    menu = QMenu()
    notification_action = menu.addAction("Notifiaciones")
    notification_action.triggered.connect(Notificaciones_emergentes_por_alerta)
    pause_action = menu.addAction("Pausar")
    pause_action.triggered.connect(Pausar_Scipt)
    exit_action = menu.addAction("Salir")
    exit_action.triggered.connect(close_application)
    tray_icon.setContextMenu(menu)
    tray_icon.show()
    app.exec_()

def Pausar_Scipt():
    global pause_threads
    if not pause_threads:
        print("Pausando")
        pause_threads = True
        thread3_event.clear()  # Desactivar el evento para pausar los hilos
    else:
        print("Reanudando")
        pause_threads = False
        thread3_event.set()  # Activar el evento para reanudar los hilos

def Notificaciones_emergentes_por_alerta():
    global Notificaciones_por_alerta
    Notificaciones_por_alerta = not Notificaciones_por_alerta
    if Notificaciones_por_alerta:
        print("Notificaciones Activadas")
        notification.notify(
            title='Notificaciones Activadas',
            message="Notificacion por alerta de Piru activado.",
            app_icon=None,
            timeout=10,
        )
        pygame.mixer.music.load(sonido)
        pygame.mixer.music.play()
    else:
        print("Notificaciones Desactivadas")
        notification.notify(
            title='Notificaciones Desactivadas',
            message="Notificacion por alerta de Piru desactivado.",
            app_icon=None,
            timeout=10,
        )
        pygame.mixer.music.load(sonido)
        pygame.mixer.music.play()

def close_application():
    global stop_threads
    print("Cerrando aplicación...")
    stop_threads = True
    thread3_event.set()
    QApplication.quit()


# Establecer la zona horaria de Argentina
try:
    argentina_tz = pytz.timezone('America/Argentina/Buenos_Aires')
except Exception as e:
    print(f"Error al establecer la zona horaria: {e}")
    argentina_tz = None

# Función para obtener la fecha y hora actuales en Argentina
def obtener_fecha_hora_arg():
    ahora_arg = datetime.now(argentina_tz)
    return ahora_arg.strftime('%Y-%m-%d %H:%M:%S')

# Función para registrar el evento en un archivo Excel
def registrar_evento(evento, nombre_archivo):
    # Inicializar COM
    pythoncom.CoInitialize()

    evento_en_succeso = obtener_fecha_hora_arg()
    nuevo_evento = pd.DataFrame([[evento_en_succeso, evento]], columns=['FechaHora', 'Evento'])
    try:
        # Leer archivo existente si lo hay
        df_existente = pd.read_excel(nombre_archivo, engine='openpyxl')
        df_actualizado = pd.concat([df_existente, nuevo_evento], ignore_index=True)
        print("Evento agregado")
    except FileNotFoundError:
        # Si no existe el archivo, crear uno nuevo
        df_actualizado = nuevo_evento
        print("Se creó el archivo")
    except Exception as e:
        print(f"Error al leer o concatenar con el archivo existente: {e}")
        df_actualizado = nuevo_evento

    #try:
        # Guardar el DataFrame actualizado en el archivo Exce'¿¿¿
    # ¡l
        #df_actualizado.to_excel(nombre_archivo, index=False, engine='openpyxl')
        #print(f"DataFrame guardado en {nombre_archivo}")
    #except Exception as e:
        #print(f"Error al guardar el DataFrame en el archivo: {e}")

    # Usar win32com para asegurarse de que se guarde el archivo correctamente en Excel
    try:
        excel = win32.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(nombre_archivo)
        workbook.Save()  # Guardar el archivo en Excel
        workbook.Close()  # Cerrar el archivo
        excel.Quit()  # Cerrar Excel
        print("Guardado forzado en Excel y sincronización asegurada.")
    except Exception as e:
        print(f"Error al forzar el guardado en Excel: {e}")

    # Finalizar COM
    pythoncom.CoUninitialize()

# Función para reemplazar un archivo en SharePoint
def reemplazar_archivo_sharepoint(driver, url_sharepoint, archivo_path):
    driver.get(url_sharepoint)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@aria-label, 'Cargar')]"))).click()

    # Esperar a que el cuadro de diálogo de carga se abra
    time.sleep(2)

    # Subir el nuevo archivo usando el cuadro de diálogo de carga
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[@name='Archivos' and @data-automationid='uploadFileCommand']//span[contains(text(), 'Archivos')]"))).click()
    time.sleep(2)
    pyautogui.write(archivo_path)
    pyautogui.press("enter")

    try:
        # Esperar a que se complete la carga
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(
            (By.XPATH, "//button[@name='Reemplazar']//span[contains(text(), 'Reemplazar')]"))).click()
        print("se cargo archivo")
        time.sleep(2)
    except:
        print("el archivo no existia")

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@aria-label, 'Cargar')]"))).click()

    # Esperar a que el cuadro de diálogo de carga se abra
    time.sleep(2)

    # Subir el nuevo archivo usando el cuadro de diálogo de carga
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[@name='Archivos' and @data-automationid='uploadFileCommand']//span[contains(text(), 'Archivos')]"))).click()
    time.sleep(2)
    pyautogui.press("enter")
    try:
        # Esperar a que se complete la carga
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(
            (By.XPATH, "//button[@name='Reemplazar']//span[contains(text(), 'Reemplazar')]"))).click()
        print("se cargo archivo")
        time.sleep(2)
    except:
        print("el archivo no existia")

def obtener_ruta_temporal(ruta_relativa):
    """Retorna la ruta al archivo empaquetado, ya sea en el entorno temporal o durante el desarrollo."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.dirname(__file__), ruta_relativa)

#-----------------------------------FUNCIONES-------------------------------------------
#Funciones_Zabbix
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

    test = "//td[text()='9s' or text()='8s' or text()='7s' or text()='6s' or text()='5s' or text()='4s' or text()='3s' or text()='2s' or text()='1s']"

    condiciones = []


    for minuto in range(30, 60):
        for segundo in range(1, 61):
            if segundo == 60:
                condiciones.append(f"text()='{minuto + 1}m 0s'")
            else:
                condiciones.append(f"text()='{minuto}m {segundo}s'")

    xpath = "//td[" + " or ".join(condiciones) + "]"

    print(xpath)
#Funcion_Iniciar_Sigma
def Iniciar_Simga():

    driver.switch_to.window(driver.window_handles[0])

    driver.get(Sigma)

    time.sleep(3)

    login_sigma_user = driver.find_element(By.ID, "j_idt8:myLogin:username")
    login_sigma_user.send_keys(Sigma_User)
    login_sigma_pass = driver.find_element(By.ID, "j_idt8:myLogin:password")
    login_sigma_pass.send_keys(Sigma_Password)
    LoginButtonSigma = driver.find_element(By.ID, "j_idt8:myLogin:loginButton")
    LoginButtonSigma.click()
    time.sleep(3)


    Home = driver.find_element(By.ID, "formHome:j_idt46:0:portalImage")
    Home.click()
    time.sleep(3)


#-----------------------------------Monitoreo-------------------------------------------
#Funcion_Monitoreo_Zabbix
def Monitoreo_de_Proxys():
    global stop_threads
    global pause_threads
    if pause_threads != False:
        thread3_event.wait()  # Pausar el hilo si el evento no está activado

    test = "//td[text()='9s' or text()='8s' or text()='7s' or text()='6s' or text()='5s' or text()='4s' or text()='3s' or text()='2s' or text()='1s']"


    driver.switch_to.window(driver.window_handles[0])


    driver.get(Zabbix)

    global xpath

    time.sleep(5)
    for _ in range(1):
        if stop_threads:
            break
        try:
            driver.get(Zabbix_Aysam)
            wait_Aysam = WebDriverWait(driver, timeout=15)
            elements = wait_Aysam.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Aysam cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Aysam', nombre_archivo=archivo_pathProxys)
                    print("se envio mail proxy aysam y se registro")
        except Exception as e:
            print(f"Pass Aysam")
        try:
            driver.get(Zabbix_BF)
            wait_BF = WebDriverWait(driver, timeout=15)
            elements = wait_BF.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Banco Formosa cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Banco Formosa', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxya BF y se registro")
        except Exception as e:
            print(f"Pass BF")
        try:
            driver.get(Zabbix_CediCBA)
            wait_cedi = WebDriverWait(driver, timeout=15)
            elements = wait_cedi.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy CediCBA cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy CediCba', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxxy cediCBA y se registro")
        except Exception as e:
            print(f"Pass CEDIDBA")
        try:
            driver.get(Zabbix_Citrusvil)
            wait_citrus = WebDriverWait(driver, timeout=15)
            print("")
            elements = wait_citrus.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Citrusvil cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Citrusvil', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por porxy Citrusvil y se registro")
        except Exception as e:
            print(f"Pass Citrusvil")
        try:
            driver.get(Zabbix_Epec2)
            wait_epec = WebDriverWait(driver, timeout=5)
            elements = wait_epec.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Epec2 cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Epec', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxy Epec y se registro")
        except Exception as e:
            print(f"Pass Epec2")
        try:
            driver.get(Zabbix_Muni)
            wait_muni = WebDriverWait(driver, timeout=5)
            elements = wait_muni.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Municipalidad cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Municipalidad', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxy Municipalidad y se registro")
        except Exception as e:
            print(f"Pass Muni")
        try:
            driver.get(Zabbix_Otek)
            wait_otek = WebDriverWait(driver, timeout=5)
            elements = wait_otek.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Otek cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Otek', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxy Otek y se registro")
        except Exception as e:
            print(f"Pass Otek")
        try:
            driver.get(Zabbix_Roela)
            wait_roela = WebDriverWait(driver, timeout=5)
            elements = wait_roela.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Roela cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy Roela', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxy Roela y se registro")
        except Exception as e:
            print(f"Pass Roela")
        try:
            driver.get(Zabbix_LR)
            wait_LR = WebDriverWait(driver, timeout=5)
            elements = wait_LR.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            for element in elements:
                if element.is_displayed():
                    texto_del_elemento = element.text
                    print("Elemento encontrado:", texto_del_elemento)
                    notification.notify(
                        title='Proxy caido',
                        message='¡Proxy Caido hace ' + texto_del_elemento + " !",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("facundo.carbajal@cedi.com.ar;mara.delgadillo@cedi.com.ar", "Proxy Caido",
                                          "Revisar Zabbix debido que un proxy Banco La Rioja cayo hace " + texto_del_elemento)
                    registrar_evento('Caido Proxy La Rioja', nombre_archivo=archivo_pathProxys)
                    print("se envio mail por proxy La Rioja y se registro")
        except Exception as e:
            print(f"Pass LR")
            time.sleep(5)
    driver.get(Zabbix)
#Funciones_Sigma
def Monitoreo_Sigma():
    if pause_threads != False:
        thread3_event.wait() # Pausar el hilo si el evento no está activado

    try:

        driver.switch_to.window(driver.window_handles[0])

        driver.get(Sigma)

        time.sleep(3)

        login_sigma_user = driver.find_element(By.ID, "j_idt8:myLogin:username")
        login_sigma_user.clear()
        login_sigma_user.send_keys(Sigma_User)
        login_sigma_pass = driver.find_element(By.ID, "j_idt8:myLogin:password")
        login_sigma_pass.send_keys(Sigma_Password)
        LoginButtonSigma = driver.find_element(By.ID, "j_idt8:myLogin:loginButton")
        LoginButtonSigma.click()
        time.sleep(3)

        Home = driver.find_element(By.ID, "formHome:j_idt46:0:portalImage")
        Home.click()
        time.sleep(3)

        driver.switch_to.window(driver.window_handles[0])

        driver.get(Sigma_Atm)

        driver.switch_to.window(driver.window_handles[-1])

        driver.close()

        driver.switch_to.window(driver.window_handles[0])

        LabelButtonSigma = driver.find_element(By.XPATH, "//span[@class='ui-icon ui-icon-triangle-1-s ui-c']")
        LabelButtonSigma.click()
        time.sleep(2)
        BancorPampaButtonSigma = driver.find_element(By.XPATH, "//li[@class='ui-selectonemenu-item ui-selectonemenu-list-item ui-corner-all' and @data-label='Banco de La Pampa' and @tabindex='-1' and @role='option' and @aria-selected='false' and @id='form:panel:grupos_1' and text()='Banco de La Pampa']")
        BancorPampaButtonSigma.click()
        time.sleep(2)

        elemento = driver.find_element(By.XPATH, "//div[@class='jqplot-point-label jqplot-series-0 jqplot-point-5']")

        texto_elemento = elemento.text
        valor = int(texto_elemento)


        if valor >= 30:
            print(f"El texto del elemento es '{valor}', que es 10 o superior. Realizando acción...")
            pygame.mixer.music.load(sonido)
            pygame.mixer.music.play()
            notification.notify(
                title='Hay mas de 10 Cajeros Fuera de Serivicio',
                message="revise nagios ante las dudas",
                app_icon=None,
                timeout=10, )
        else:
            print(f"El texto del elemento es '{valor}', que es menor que 10. No se realiza ninguna acción.")

        time.sleep(5)
    except:
        print("no se logro iniciar a Sigma")
        pygame.mixer.music.load(sonido)
        pygame.mixer.music.play()
        notification.notify(
            title='no se logro iniciar a Sigma',
            message="chequear la pagina despues del recorrido",
            app_icon=None,
            timeout=10, )



#Funcion_Jira_Backlog111
def Monitoreo_Jira_Backlog():
    global numero_de_tickets
    if pause_threads != False:
        thread3_event.wait() # Pausar el hilo si el evento no está activado


    driver.switch_to.window(driver.window_handles[0])

    driver.get(Jira_Backlog)

    for _ in range(1):
        if stop_threads:
            break
        try:
            element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//span[@class="filter-render-view-number"]'))
            )
            texto_del_elemento = element.text
            numero_de_tickets = int(texto_del_elemento)

            if 1 <= numero_de_tickets <= 100:
                print(f"Elemento encontrado con {numero_de_tickets} tickets.")
                notification.notify(
                    title='Hay Tickets sin Asignar',
                    message=f"Hay {numero_de_tickets} Ticket(s) sin Asignar",
                    app_icon=None,
                    timeout=60,
                )
                pygame.mixer.music.load(sonido)
                pygame.mixer.music.play()
            else:
                print(f"Elemento encontrado con {numero_de_tickets} tickets, no se requiere notificación.")

        except Exception as e:
            print(f"No hay tickets")


    time.sleep(10)
#Funcion_MMG
# Función que verifica los tickets sin asignar
def Monitoreo_MMG():
    if pause_threads:
        thread3_event.wait()

    driver.switch_to.window(driver.window_handles[0])

    driver.get(MMG_Login)
    time.sleep(3)

    try:
        login_mmg_user = driver.find_element(By.XPATH, "//input[@ng-model='user_name']")
        login_mmg_user.send_keys(MMG_User)
        login_mmg_pass = driver.find_element(By.XPATH, "//input[@ng-model='user_password']")
        login_mmg_pass.send_keys(MMG_Pass)
        LoginButtonMMG = driver.find_element(By.XPATH, "//button[@name='login']")
        LoginButtonMMG.click()
        time.sleep(3)
        print("Se inició con éxito")
    except Exception as e:
        print("No se logró iniciar MMG")

        pyautogui.moveTo(1335, 141)
        # pyautogui.moveTo(1338, 113)
        time.sleep(5)
        pyautogui.click(1335, 141)
        # pyautogui.click(1338,113)

    driver.get(MMG_Tickets)
    time.sleep(4)

    try:
        # Encuentra el elemento de la columna

        pyautogui.moveTo(367, 239)
        time.sleep(3)


        columna_mmg = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//i[@class='icon-ellipsis-vertical col-menu list_header_context list-column-icon']")))
        columna_mmg.click()

        print("Se hizo clic en el menú de columna")

        export_mmg = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//div[@item_id='d1ad2f010a0a0b3e005c8b7fbd7c4e28' and @data-context-menu-label='Export']")))

        export_mmg.click()
        print("Se encontró el elemento export")


        export_mmg.click()

        excel_mmg = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//div[@item_id='f13f0041473012003db6d7527c9a71f0' and text()='Excel (.xlsx)']")))

        print("Se encontró el elemento excel")

        excel_mmg.click()

        download_mmg = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//button[@id='download_button' and text()='Download']")))

        print("Se encontró el botón de descarga")

        download_mmg.click()

    except Exception as e:
        print(f"Error al verificar tickets o descargar archivo: {e}")
    time.sleep(5)

def Monitoreo_MMG_2():
    global  numero_de_tickets_mmg, MMG_User,MMG_Pass, MMG_Tickets
    if pause_threads:
        thread3_event.wait()

    driver.switch_to.window(driver.window_handles[0])
    driver.execute_script("document.body.style.zoom='75%'")
    driver.get(MMG_Login)
    time.sleep(3)
    driver.execute_script("document.body.style.zoom='75%'")

    pyautogui.moveTo(1335,141)
    #pyautogui.moveTo(1338, 113)
    time.sleep(5)
    pyautogui.click(1335,141)
    #pyautogui.click(1338,113)


    try:
        login_mmg_user = driver.find_element(By.XPATH, "//input[@ng-model='user_name']")
        login_mmg_user.send_keys(MMG_User)
        login_mmg_pass = driver.find_element(By.XPATH, "//input[@ng-model='user_password']")
        login_mmg_pass.send_keys(MMG_Pass)
        LoginButtonMMG = driver.find_element(By.XPATH, "//button[@name='login']")
        LoginButtonMMG.click()
        time.sleep(3)
        print("Se inició con éxito")
    except Exception as e:
        print("No se logró iniciar MMG")

    warnings.filterwarnings("ignore",category=UserWarning, module="openpyxl")

    driver.get(MMG_Tickets)

    time.sleep(3)


    pyautogui.moveTo(367,239)
    time.sleep(3)
    pyautogui.click(367,239)

    #pyautogui.moveTo(486,281)
    #time.sleep(3)
    #pyautogui.click(486,281)

    pyautogui.moveTo(421,393)
    time.sleep(3)
    pyautogui.click(421,393)

    #pyautogui.moveTo(572,491)
    #time.sleep(3)
    #pyautogui.click(572,491)

    pyautogui.moveTo(569,398)
    time.sleep(3)
    pyautogui.click(569,398)

    #pyautogui.moveTo(732,494)
    #time.sleep(3)
    #pyautogui.click(732,494)

    pyautogui.moveTo(769,467)
    time.sleep(5)
    pyautogui.click(769,467)

    #pyautogui.moveTo(803,457)
    #time.sleep(5)
    #pyautogui.click(803,457)

    time.sleep(6)

    global Incidentes_MMG, df
    try:
        df = pd.read_excel(Incidentes_MMG)
        print("Archivo leido con exito")
    except:
        "problema al leer el archivo"
    try:
        for i in range(len(df)):
            numero = df.loc[i, "Number"]
            asignado = df.loc[i, "Assigned to"]
            asignado = str(asignado)
            if asignado == "nan":
                notification.notify(
                    title='Ticket ' + numero + " sin asignar",
                    message='Ticket ' + numero + " sin asignar",
                    app_icon=None,
                    timeout=120
                )
                pygame.mixer.music.load(sonido)
                pygame.mixer.music.play()
                print('Ticket ' + numero + " sin asignar")
                numero_de_tickets_mmg=numero_de_tickets_mmg+1
    except:
        print("Problemas al leer el archivo")
        numero_de_tickets_mmg=0
    time.sleep(3)
    try:
        os.remove(Incidentes_MMG)
        print("Archivo de Incidentes Borrado")
    except:
        print("No se pudo borrar el archivo")
    time.sleep(2)

#Funciones_Piru
def Monitoreo_de_Piru():
    if pause_threads != False:
        thread3_event.wait() # Pausar el hilo si el evento no está activado

    global Notificaciones_por_alerta

    driver.switch_to.window(driver.window_handles[0])

    driver.get(Piru)

    time.sleep(10)


    for _ in range(15):
        if stop_threads:
            break
        try:
            wait = WebDriverWait(driver, 30 )
            elemento_clicleable=wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='No']")))
            elemento_clicleable.click()
            print("encontrado")
            if Notificaciones_por_alerta != False:
                notification.notify(
                     title='NUEVA ALERTA',
                     message="Nueva alerta cedi",
                     app_icon=None,
                     timeout=10, )
                pygame.mixer.music.load(sonido)
                pygame.mixer.music.play()
            time.sleep(20)
        except Exception as e:
            print(f"No se encontro elemento aun.")
            time.sleep(15)
#Funcion_Nagioscedi
hth_caido = False
hth_t_caido = False
global hth_trasacciones_caido, hth_trasacciones_okey,  element_t,element_h ,hth_okey, wait
#Funcion_Nagios
def Monitoreo_Nagios():
    global hth_caido
    global hth_t_caido
    global inicio_nagios
    if pause_threads:
        thread3_event.wait()  # Pausar el hilo si el evento no está activado

    driver.switch_to.window(driver.window_handles[0])

    if inicio_nagios==False:
        driver.get(Nagios_Front)
        pyautogui.write(Nagios_User)
        pyautogui.press("tab")
        pyautogui.write(Nagios_Pass)
        pyautogui.press("enter")
        inicio_nagios=True


    try:
        login_nagios_user = driver.find_element(By.ID, "usernameBox")
        login_nagios_user.send_keys(Nagios_User)
        login_nagios_pass = driver.find_element(By.ID, "passwordBox")
        login_nagios_pass.send_keys(Nagios_Pass)
        LoginButtonNagios = driver.find_element(By.ID, "loginButton")
        LoginButtonNagios.click()
        time.sleep(3)
    except Exception as e:
        pyautogui.write(Nagios_User)
        pyautogui.press("tab")
        pyautogui.write(Nagios_Pass)
        pyautogui.press("enter")
        print("Nagios ya iniciado")

    driver.get(Nagios_Front)

    for _ in range(1):
        if stop_threads:
            break
        try:
            wait = WebDriverWait(driver, timeout=10)
            hth_transacciones = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//img[@src='/nagvis/userfiles/images/iconsets/circulo_22_critical.png' and @alt='service-SR001K12033-Check HTH transacciones']")))
            for element in hth_transacciones:
                if element.is_displayed() and not hth_t_caido:
                    notification.notify(
                        title='HTH TRANSACCIONES CAIDO',
                        message="HTH TRANSACCIONES CAIDO",
                        app_icon=None,
                        timeout=120
                    )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("noc.cedi@cedi.com.ar",
                                          "HTH TRANSACCIONES ESTA CAIDO",
                                          "HTH TRANSACCIONES ESTA CAIDO")
                    registrar_evento('HTH TRANSACCIONES CAIDO',nombre_archivo=archivo_path)
                    print("Se envió el mail y se registró el evento")
                    notification.notify(
                        title='Se registro evento',
                        message="Se procede a guardarlo en el sharepoint",
                        app_icon=None,
                        timeout=120
                    )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    hth_t_caido = True
        except Exception as e:
            print(f"HTH TRANSACCIONES ESTA OKEY")
        try:
            wait = WebDriverWait(driver, timeout=10)
            hth_transacciones = wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                                "//img[@id='e73d84-icon' and @class='icon' and @src='/nagvis/userfiles/images/iconsets/circulo_22_ok.png' and @alt='service-SR001K12033-Check HTH transacciones']")))
            for element in hth_transacciones:
                if element.is_displayed() and hth_t_caido == True:
                    print("HTH_T se restablecio")
                    enviar_correo_outlook("noc.cedi@cedi.com.ar",
                                          "HTH_T se restablecio",
                                          "HTH_T se restablecio")
                    registrar_evento('HTH_T se restablecio', nombre_archivo=archivo_path)
                    print("Se registró el evento HTH_T")
                    notification.notify(
                        title='Se registro evento',
                        message="Se procede a guardarlo en el sharepoint",
                        app_icon=None,
                        timeout=120
                    )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    hth_t_caido = False
        except Exception as e:
            print(f"HTH TRANSACCIONES aun sigue caido")
        try:
            wait = WebDriverWait(driver, timeout=10)
            hth_host = wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                       "//img[@id='2489da-icon' and @src='/nagvis/userfiles/images/iconsets/circulo_22_critical.png']")))
            for element_h in hth_host:
                if element_h.is_displayed and hth_caido == False:
                    notification.notify(
                        title='HTH HOST CAIDO',
                        message="HTH HOST CAIDO",
                        app_icon=None,
                        timeout=120, )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    enviar_correo_outlook("noc.cedi@cedi.com.ar",
                                          "HTH CAIDO",
                                          "HTH CAIDO")
                    print("se envió mail")
                    registrar_evento('HTH CAIDO', nombre_archivo=archivo_path)
                    print("Se registró el evento HTH CAIDO")
                    notification.notify(
                        title='Se registro evento',
                        message="Se procede a guardarlo en el sharepoint",
                        app_icon=None,
                        timeout=120
                    )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    hth_caido = True
        except Exception as e:
            print(f"HTH ESTA OKEY")
        try:
            wait = WebDriverWait(driver, timeout=10)
            hth_host = wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                       "//img[@id='2489da-icon' and @class='icon' and @src='/nagvis/userfiles/images/iconsets/circulo_22_up.png' and @alt='host-host to host BPI']")))
            for element_h in hth_host:
                if element_h.is_displayed and hth_caido == True:
                    print("Se restablecio HTH")
                    registrar_evento('HTH Restablecio', nombre_archivo=archivo_path)
                    enviar_correo_outlook("noc.cedi@cedi.com.ar",
                                          "HTH se restablecio",
                                          "HTH se restablecio")
                    print("Se registró el evento HTH")
                    notification.notify(
                        title='Se registro evento',
                        message="Se procede a guardarlo en el sharepoint",
                        app_icon=None,
                        timeout=120
                    )
                    pygame.mixer.music.load(sonido)
                    pygame.mixer.music.play()
                    hth_caido = False
        except Exception as e:
            print(f"HTH aun sigue caido")



    time.sleep(3)

def iniciar_tkinter_en_hilo():
    global stop_threads
    hilo_tkinter= threading.Thread(target=iniciar_tkinter)
    hilo_tkinter.start()



def Correr_thread():
    obtener_ruta_temporal(ruta_eventos)
    obtener_ruta_temporal(ruta_sonido)
    global stop_threads,numero_de_tickets,numero_de_tickets_mmg
    stop_threads = False
    driver.execute_script("window.open('');")
    pygame.mixer.music.load(sonido)
    Asignacion_de_Credenciales_Keepass()
    Asignacion_de_URLS_Keepass()
    iniciar_tkinter_en_hilo()
    pygame.mixer.music.play()
    Iniciar_Zabbix_Proxys()
    notification.notify(
        title='Se toma el Control durante 2 min',
        message="Aguarde",
        app_icon=None,
        timeout=10,)
    while not stop_threads:
        Monitoreo_de_Proxys()
        print("PASE A SIGMA")
        time.sleep(3)
        if stop_threads:
            break
        Monitoreo_Sigma()
        print("PASE A JIRA")
        time.sleep(3)
        if stop_threads:
            break
        Monitoreo_Jira_Backlog()
        actualizar_tickets()
        print("PASE A NAGIOS")
        time.sleep(3)
        if stop_threads:
            break
        Monitoreo_Nagios()
        print("PASE A PIRU")
        if stop_threads:
            break
        pygame.mixer.music.load(sonido)
        pygame.mixer.music.play()
        notification.notify(
            title='Puede retomar el control',
            message="Gracias por la espera",
            app_icon=None,
            timeout=10, )
        Monitoreo_de_Piru()
        pygame.mixer.music.load(sonido)
        pygame.mixer.music.play()
        notification.notify(
            title='Se toma el Control durante 2 min',
            message="Aguarde",
            app_icon=None,
            timeout=10, )
        print("REINICIO")
        time.sleep(3)
        if stop_threads:
            break

if __name__ == "__main__":
    driver = inicializar_driver()
    inicio_nagios = False
    # ------------------cedi-----------------------URlS----------------------------------
    Piru = ""
    Zabbix = ""
    Zabbix_Aysam= ""
    Zabbix_BF=""
    Zabbix_CediCBA=""
    Zabbix_Citrusvil=""
    Zabbix_Epec2=""
    Zabbix_Muni=""
    Zabbix_Otek=""
    Zabbix_Roela=""
    Zabbix_LR=""
    Nagios_Front =""
    Nagios=""
    Sigma = ""
    Sigma_Atm = ""
    Jira_Backlog = ""
    MMG_Login = ""
    MMG_Tickets =""
    url_sharepoint=""
    # -----------------------------------------Rutas de Archivos----------------------------------

    #Ruta de DTB Keepass
    user_folder= os.getenv('USERPROFILE')
    ruta_kdbx = os.path.join(user_folder,'OneDrive - CEDI TECH Consulting','Monitoreo 2024','Caperta para Script','DataBaseKeepass','Credenciales.kdbx')
    kp=PyKeePass(ruta_kdbx,password="AsistenteOctubre2024")


    # Ruta de eventos.xlsx
    ruta_eventos = obtener_ruta_temporal(os.path.join(user_folder,'OneDrive - CEDI TECH Consulting','Monitoreo 2024','Caperta para Script','Eventos','eventos.xlsx'))
    archivo_path = ruta_eventos

    # Ruta de Registro_Proxys.xlsx
    ruta_proxys = obtener_ruta_temporal(os.path.join(user_folder,'OneDrive - CEDI TECH Consulting','Monitoreo 2024','Caperta para Script','Eventos','Registro_Proxys.xlsx'))
    archivo_pathProxys = ruta_proxys

    # Ruta de Credenciales.xlsx
    ruta_credenciales = obtener_ruta_temporal(os.path.join("Elementos", "Credenciales", "Credenciales.xlsx"))
    credenciales_excel = ruta_credenciales

    # Ruta de Urls.xlsx
    ruta_urls = obtener_ruta_temporal(os.path.join("Elementos", "Urls", "Urls.xlsx"))
    Urls = ruta_urls

    #Ruta de Incidentes MMG
    ruta_incidentes = obtener_ruta_temporal(os.path.join(user_folder,"Downloads","incident.xlsx"))
    Incidentes_MMG=ruta_incidentes

    ruta_sonido = obtener_ruta_temporal(os.path.join("Elementos", "Notificacion.mp3"))
    # -----------------------------------------CREDENCIALES----------------------------------
    Sigma_User = ""
    Sigma_Password = ""
    MMG_User = ""
    MMG_Pass = ""
    Nagios_User = ""
    Nagios_Pass = ""
    numero_de_tickets=0
    numero_de_tickets_mmg=0

    thread3 = Thread(target=Correr_thread)

    thread3.start()

    tray_thread = Thread(target=create_tray_icon)
    tray_thread.start()