
# Coding in times of COVID-19.
# Realizado con la intención de optimizar procesos de revisión estado de licencias médicas
# por personal del HGGB Concepción.
# Hecho por Jonathan Friz B. jfriz[@]protonmail.com
# reemplaza Licencias_beta.xlsx por nombre de tu archiivo .xlsx)
import pandas as pd
import json, xlsxwriter, time, random, requests
from termcolor import colored
from xlsxwriter.utility import xl_rowcol_to_cell as to_cell
from termcolor import colored

print()
print (colored('Coding in times of COVID-19.', 'green', attrs=['underline']))
print()
print(colored('Work smarter', 'red', attrs=['bold']), 'and harder!')
print()

start_time = time.time()
print('Hora Inicio', time.asctime())
print()
start_time = time.time()

df = pd.read_excel("Licencias_beta.xlsx")  #Reemplaza nombre de tu arhivo aquí.

# ----------------------------------------------------------
# Crear el archivo fuera del loop

workbook = xlsxwriter.Workbook('Licencias_revisadas.xlsx')   
worksheet = workbook.add_worksheet()                    
# ----------------------------------------------------------
#celda = (df.iloc[1,9]) lectura por celda  # es una alternativa, sin embargo no es la mejor
# Obtener todos los links de una vez

# lista_links = df['LINK'].to_list() hasta version 2.4 luego en 3.0 se lee directamente desde excel desde la base y se generan los links

# ----------------------------------------------------------
# se setea url base de request del sitio

url = "https://bi769qi4r4.execute-api.us-east-2.amazonaws.com/prod/get-estado-mlm?lmid="

# folio_str = str(folio)

df = pd.read_excel("Licencias_beta.xlsx")    #Reemplaza nombre de tu arhivo aquí.
run_ori = df['Rut'].tolist()
folio_str = (df['Folio'].tolist())
rows = df[['Rut', 'Folio']]

# print(rows.info()) 
# Si rows fuera una lista, iteramos como 'for row in rows':
# rows = [1,2,3,4]

# en pandas para iterar debemos ir a traves de iterrows()
lista_links = []
for i, row in rows.iterrows():
    rut = row['Rut']
    folio = row['Folio']
    if rut <= 9999999:
        rut_ceros = '{:0>10}'.format(rut) # esto me agrega 3 ceros antes del numero
    else:
        rut_ceros = '{:0>10}'.format(rut)   # esto me agrega 2 ceros
    print(rut_ceros, folio)
    folio_str = str(folio)
    link = url+rut_ceros+"_"+folio_str
    lista_links.append(link)
    #print(link)

# ----------------------------------------------
# Iteramos hasta la ultimo link de la lista
#
for i in range(0, len(lista_links)):
    try:
        link = requests.get(lista_links[i])
        json_web = json.loads(link.text)
        estado_licencia = json_web['body']['instancias_v2'][0]['scripts']['script_h4']
        pago_licencia = json_web['body']['instancias_v2'][0]['scripts']['script_hito']
        nombre_licencia = json_web['body']['nom_trab']
        run_licencia = json_web['body']['rut_trab']
        folio_licencia = json_web['body']['folio'].strip()[2:11]

        print(i+1, nombre_licencia, folio_licencia, estado_licencia, pago_licencia)

        # Añadir la info al nuevo xlsx
        worksheet.write(to_cell(0, 0), "Nombre")
        worksheet.write(to_cell(0, 1), "RUN")
        worksheet.write(to_cell(0, 2), "Folio")
        worksheet.write(to_cell(0, 3), "Estado")
        worksheet.write(to_cell(0, 4), "Derecho a pago")
        worksheet.write(to_cell(i+1, 0), nombre_licencia)
        worksheet.write(to_cell(i+1, 1), run_licencia)
        worksheet.write(to_cell(i+1, 2), folio_licencia)
        worksheet.write(to_cell(i+1, 3), estado_licencia)
        worksheet.write(to_cell(i+1, 4), pago_licencia)

        # random.uniform genera numero aleatorio de segundos de espera entre 3 y 6 segundos para dejar descansar el servidor remoto entre request y request

        wait_time = (random.uniform(2.8, 4))
        print(f'{wait_time:.4f}', 'Segundos de espera aleatorios entre requerimientos al servidor.')
        print()
        time.sleep(wait_time)
        
    except KeyError: 
        worksheet.write(to_cell(i+1, 0), 'Folio Errado')
        worksheet.write(to_cell(i+1, 1), 'Folio Errado)
        worksheet.write(to_cell(i+1, 2), 'Folio Errado')
        worksheet.write(to_cell(i+1, 3), 'Folio Errado')
        worksheet.write(to_cell(i+1, 4), 'Folio Errado')
        print(colored('Folio Errado - Verificar', 'red', attrs=['bold']))
        pass
    except TypeError:
        worksheet.write(to_cell(i+1, 0), 'Folio no encontrado Revisar a mano')
        worksheet.write(to_cell(i+1, 1), 'Folio no encontrado Revisar a mano')
        worksheet.write(to_cell(i+1, 2), 'Folio no encontrado Revisar a mano')
        worksheet.write(to_cell(i+1, 3), 'Folio no encontrado Revisar a mano')
        worksheet.write(to_cell(i+1, 4), 'Folio no encontrado Revisar a mano')
        print(colored('Folio no encontrado - Verificar', 'red', attrs=['bold']))
        pass
# ----------------------------------------------------------
# Terminamos de escribir y cerramos el documento
workbook.close()

print('Terminado!')
stop_time = time.time()
dt = stop_time - start_time 
print()
print('Hora Finalizado', time.asctime())
print()
dt_min = (dt/60)
print()
print('Acabas de ahorrar',  (f'{dt_min:.3f}'), "minutos, para hacer mejor tu trabajo.")
print()

# ----------------------------------------------------------
# The nerd part
# ----------------------------------------------------------
print(colored("Never send a human to do a machine's job!", "red", attrs=["bold"]))

    
#   ╦┌─┐┌┐┌┌─┐┌┬┐┬ ┬┌─┐┌┐┌   ╔═╗┬─┐┬┌─┐  ╔╗  
#   ║│ ││││├─┤ │ ├─┤├─┤│││   ╠╣ ├┬┘│┌─┘  ╠╩╗ 
#  ╚╝└─┘┘└┘┴ ┴ ┴ ┴ ┴┴ ┴┘└┘   ╚  ┴└─┴└─┘  ╚═╝o

#
# Mejoras desde versión 1. Se maneja errores de Key Error con excepción para que no se detenga script, 
# Al igual con TypeError.
# Se agregan colores y formatos para visualización en consola
# Se agrega tiempos de esperas ya no fijos, sino aleatorios dentro de un rango para reposo del script 
# y evitar bloqueos y sobrecargas del sitio web. 
# Se agrega hora de inicio y final con marca de tiempo y se hace diferencia de la duración del proceso completo.
# Se acorta a escribar en planilla de salida ahora se registrara sin el primer digito ni el guion para mejor compración, 
# el sitio web lo entrega con este digito.
# ----------------------------------------------------------
# Version 3.0 Se generan los links en python directo de lista de RUT y Folios.
# Se genera validación de ruts mayores y menores a 10 millones.
# ----------------------------------------------------------
# 3.1 Se agregan índices o headers a la planilla de salida
