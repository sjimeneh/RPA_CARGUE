import os
import glob
import asyncio
from pyppeteer import launch
from dotenv import dotenv_values
import shutil
import datetime



env_var = dotenv_values(".env")

def validar_archivos_por_cargar():
    #Buscar todos los archivos que terminen en .xlsx
    carpeta = f"{env_var['SERVER_NAME']}{env_var['SERVER_NAME_FOLDER_LOAD']}"
    extension = env_var["EXTENSION_ARCHIVO"]

    archivos = glob.glob(os.path.join(carpeta, f'*{extension}'))

    if bool(archivos):
        # Crear una lista de tuplas (archivo, fecha_de_creacion)
        archivos_con_fechas = [(archivo, os.path.getctime(archivo)) for archivo in archivos]

        # Ordenar la lista de archivos por fecha de creación de mayor a menor
        archivos_ordenados = sorted(archivos_con_fechas, key=lambda x: x[1], reverse=True)

        # Obtener solo la lista de archivos ordenados
        archivos_ordenados = [archivo[0] for archivo in archivos_ordenados]

    
        return True,archivos_ordenados
    else:
        return False,[]

# Función para obtener la fecha de creación de un archivo
def obtener_fecha():
    # Obtener la fecha y hora actual
    fecha_hora_actual = datetime.datetime.now()

    # Formatear la fecha y hora en el formato deseado
    fecha_hora_formateada = fecha_hora_actual.strftime('%Y-%m-%d %H:%M:%S')

    return str(fecha_hora_formateada)

async def existe_elemento(page,tipo,selector):
    try:
        if(tipo == "id"):
            element = await page.querySelector(selector)
            if element:
                return True
            else:
                print("No se ha encontrado el elemento: "+selector)
                return False
    except Exception as e:
        print("No se ha encontrado el elemento: "+selector)
        return False

async def logueo_ows(page):
    try:
        elemento_user = env_var["OWS_SELECTOR_ID_USER"]
        elemento_password = env_var["OWS_SELECTOR_ID_PASSWORD"]
        usuario_ows = env_var["OWS_USER"]
        password_ows = env_var["OWS_PASSWORD"]
        boton_logueo_ows = env_var["OWS_SELECTOR_BOTON_LOGUEO"]
        if(await existe_elemento(page,"id",elemento_user) == True):

            # llenar los campos de inicio de sesión
            await page.type(elemento_user, usuario_ows)
            await page.type(elemento_password, password_ows)

            # Hacer clic en el botón de inicio de sesión
            await page.click(boton_logueo_ows)
    except Exception as e:
        await print(f"ERROR - {e} - [POSIBLE ERROR EN LA LECTURA DE VARIABLES .ENV]")

async def cargar_archivo_ows(page,ruta_archivo):
    try:
        await asyncio.sleep(5)
        boton_ok_selected = env_var["OWS_SELECTOR_ID_BOTON_OK"]
        boton_import = env_var["OWS_SELECTOR_ID_BOTON_IMPORT"]
        

        # Seleccionar el elemento file
        input_file = await page.querySelector('input[type="file"]')

        # Cargar el archivo en el elemento file
        await input_file.uploadFile(ruta_archivo)

        await asyncio.sleep(3)

        if(await existe_elemento(page,"id",boton_ok_selected)):
            await page.click(boton_ok_selected)

        # Presionar Click en el elemento Import
        await page.click(boton_import)
        return True
    except Exception as e:
        print("[ERROR] - [LA CARGA A OWS NO SE HA COMPLETADO]")
        print(f"[MENSAJE] - [{str(e)}]")
        return False

async def validar_carga(page):
    carga_exitosa = True
    try:
        ows_text_status_fail = env_var["OWS_TEXT_STATUS_FAIL"]
        # Utilizar page.evaluate para obtener el contenido del elemento
        respuesta_status = await page.evaluate('() => { return document.querySelector("#ext-gen32").textContent; }')
        respuesta_errorMessage = await page.evaluate('() => { return document.querySelector("#ext-gen40").textContent; }')

        if(respuesta_status == ows_text_status_fail):
            carga_exitosa = False
    except Exception as e:
        print("[ERROR] - [LA VALIDCION DE CARGA HA FALLADO]")
        print(f"[MENSAJE] - [{str(e)}]")
        carga_exitosa = False
        respuesta_errorMessage = "ARCHIVO CORRUPTO"
    finally:
        return carga_exitosa,respuesta_errorMessage

async def crear_archivo_error_carga(nombre_archivo,texto):
    try:
        ruta_archivo = f"{env_var['SERVER_NAME']}{env_var['SERVER_NAME_FOLDER_FAIL_LOAD']}\\{nombre_archivo}"
        # Abre el archivo en modo de escritura ('w')
        with open(ruta_archivo, 'w') as archivo:
            archivo.write(texto)
    except Exception as e:
        print("[ERROR] - [NO SE HA PODIDO CREAR EL ARCHIVO ERRORES CARGUE]")
        print(f"[MENSAJE] - [{str(e)}]")

async def mover_archivo_resumen(nombre_archivo):
    try:
        ruta_origen = f"{env_var['SERVER_NAME']}{env_var['SERVER_NAME_FOLDER_FAIL_LOAD']}\\{nombre_archivo}"
        ruta_destino = f"{env_var['SERVER_NAME']}{env_var['SERVER_NAME_FOLDER_SUMMARY_LOAD']}\\{nombre_archivo}"
        # Mover el archivo
        shutil.move(ruta_origen, ruta_destino)
    except Exception as e:
        print("[ERROR] - [EL ARCHIVO NO SE PUDO MOVER A LA RUTA DESTINO]")
        print(f"[MENSAJE] - [{str(e)}]")

async def remover_archivo(nombre_archivo):
    try:
        ruta_archivo = f"{env_var['SERVER_NAME']}{env_var['SERVER_NAME_FOLDER_LOAD']}\\{nombre_archivo}"
        if os.path.exists(ruta_archivo):
            os.remove(ruta_archivo)
    except Exception as e:
        print(f"[ERROR] - [EL ARCHIVO NO SE PUDO ELIMINAR]")
        print(f"[MENSJAE] - [{str(e)}]")

async def navegador_ows(browser):
    url_ows = env_var["OWS_URL"]
    numero_iteracion = 0
    while(True):
        #Aumento el numero de Iteraciones en 1
        numero_iteracion += 1
        hay_archivos,lista_archivos = validar_archivos_por_cargar()

        if(hay_archivos):

            #Abrir una nueva pestaña
            page = await browser.newPage() 

            # Establecer las dimensiones de la ventana igual a las dimensiones de la pantalla
            await page.setViewport({'width': 1080, 'height': 1080})

            try:
                #Ir a la url de OWS
                await page.goto(url_ows)

                #Si el elemento por id username está visible, realice el logueo
                await logueo_ows(page)

                #Recorrer cada uno de los archivos por cargar
                for archivo in lista_archivos:

                    if(await cargar_archivo_ows(page,archivo)):
                        await asyncio.sleep(10)

                        #Obtener solo el nombre del archivo
                        nombre_archivo = os.path.basename(archivo)

                        #Comprobar los estados de la respuesta
                        carga_exitosa, mensajeError =  await validar_carga(page)

                        if(carga_exitosa):
                            print("[EXITOSO] [LA CARGA A OWS HA SIDO SATISFACTORIA]")
                            await crear_archivo_error_carga(nombre_archivo,"Carga Exitosa")
                            await mover_archivo_resumen(nombre_archivo)
                        else:
                            print("[ERROR] [LA CARGA A OWS HA FALLADO]")
                            print(f"[MENSAJE] [{mensajeError}]")
                            await crear_archivo_error_carga(nombre_archivo,f"Fail\n\n{mensajeError}")
                            await remover_archivo(nombre_archivo)

                        print("\n\n")

                        #Recargar nuevamente la pagina
                        await page.goto(url_ows)

                    else:
                        print("[ERROR] [LA CARGA A OWS HA FALLADO]")
                        print(f"[MENSAJE] [{mensajeError}]")
                        await crear_archivo_error_carga(nombre_archivo,f"Fail\n\n{mensajeError}")
                        await remover_archivo(nombre_archivo)
                        await page.goto(url_ows)

            except Exception as e:
                print(f"ERROR - {e}")
            finally:
                print("FIN DEL PROGRAMA")
                await page.close()
        else:
            print(f"[{obtener_fecha()}] [NO HAY ARCHIVOS PARA CARGAR]")
        
        await asyncio.sleep(5)

        #REINCIAR EL PROGRAMA
        if(numero_iteracion >= 4000):
            #Reiniciando el programa
            await browser.close()
            break

async def main():
    while(True):
        try:
            # Lanzar el navegador
            browser = await launch(
                ignoreHTTPSErrors=True,
                # executablePath='/root/.cache/puppeteer/chrome/linux-114.0.5735.90/chrome-linux64/chrome',
                args=['--no-sandbox',
                '--proxy-server=http://10.158.122.48:8080',
                '--disable-infobars',  # Deshabilitar la barra de información (como las cookies)
                '--disable-notifications',  # Deshabilitar las notificaciones del navegador
                '--disable-extensions',
                '--disable-gpu',
                '--no-sandbox',
                '--disable-dev-shm-usage',
                    ], 
                    options={
                        'headless': False,  # Habilita la visualización del navegador
                        'userDataDir': "temporal_data" #Cambiar la ruta de los datos temporales
                            }
                            )
            await navegador_ows(browser)
        except Exception as e:
            print("[ERROR] - [EL HILO PRINCIPAL HA FALLADO]")
            print(f"[MENSAJE] - [{str(e)}]")
        finally:
            await browser.close()
            print("60 SEGUNDOS PARA EL RESTABLECIMIENTO DE CARGUE")
            await asyncio.sleep(60)

# Ejecutar la función asincrónica
if __name__ == "__main__":
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(main())