import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
import pyfiglet
from colorama import Fore
from requests.exceptions import HTTPError
import re
from collections import Counter
import os
import xlsxwriter
#----------------------------URLS----------------------------#

seccion={
    1:'https://www.el-carabobeno.com/secciones/noticias/nacional/',
    2:'https://www.el-carabobeno.com/secciones/noticias/internacional/',
    3:'https://www.el-carabobeno.com/secciones/deportes/',
    4:'https://www.el-carabobeno.com/secciones/noticias/valencia/',
    5:'https://www.el-carabobeno.com/secciones/universidad/',
}

#--------------------------STOPWORDS--------------------------#
STOPWORDS = set([
    'de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'un', 'una', 'por', 'con', 'su', 
    'para', 'del', 'las', 'al', 'lo', 'se', 'no', 'como', 'más', 'o', 'pero', 'sus', 
    'le', 'ya', 'hasta', 'si', 'sin', 'sobre', 'este', 'esta', 'estos', 'estas', 'fue', 'ser'
])
#-----------------------------DICCIONARIOS Y LISTAS RELEVANTES-----------------------------------#

nombres={
    1:'Nacional',
    2:'Internacional',
    3:'Deportes',
    4:'Gran Valencia',
    5:'Universidad',
}
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "es-ES,es;q=0.9",
    "Referer": "https://www.google.com/"
}
#Aqui se colocaran las palabras que el usuario desea revisar a la hora de scrapear las noticias
keywords=[]

# Genera un archivo Excel con formato profesional (Encabezados azules, columnas ajustadas)
def exportarExcelBonito(df, nombre_base):
    
    nombre_archivo = f"{nombre_base}.xlsx"
    
    try:
        writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
        
        # Convertimos el dataframe a Excel
        sheet_name = 'Noticias'
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Obtenemos los objetos para trabajar el formato
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # Formato para encabezados (Azul, Negrita, Texto Blanco)
        formato_header = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#0070C0', 
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center'
        })
        # Formato para el cuerpo del texto
        formato_texto = workbook.add_format({
            'text_wrap': True, 
            'valign': 'top',
            'border': 1
        })
        
        # Formato para Links (opcional, azul y subrayado)
        formato_link = workbook.add_format({
            'font_color': 'blue',
            'underline': 1,
            'text_wrap': True,
            'valign': 'top',
             'border': 1
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, formato_header)
   
        worksheet.set_column('A:A', 50, formato_texto) 
        worksheet.set_column('B:B', 20, formato_texto) 
        worksheet.set_column('C:C', 40, formato_link)

        writer.close()
        print(Fore.GREEN + f"--> EXCEL GENERADO EXITOSAMENTE: {nombre_archivo}")
        
    except Exception as e:
        print(Fore.RED + f"Error generando el Excel: {e}")
        
        
#Menu se encarga de poder visualizar las opciones que requerimos al seleccionar la seccion que deseamos scrapear
def menu():
    
    continuar=True
    while continuar:
        
        # Limpiar pantalla (funciona en windows y linux)
       # os.system('cls' if os.name == 'nt' else 'clear')
        
        titulo = pyfiglet.figlet_format('Scraper Carabobeno')
        print(Fore.GREEN + titulo)
        print(Fore.WHITE + "Aqui podras visualizar las distintas noticias\n")
        try:
            sec = int(input("Que deseas visualizar:\n1. Nacional\n2. Internacional\n3. Deportes\n4. Gran Valencia\n5. Universidad\n6.Opcion por actualizar(No disponible)\n\nOpción: "))
            
            if sec in seccion:
                print(Fore.YELLOW + "\nConectando con la pagina, espera un momento...\n")
                noticias=analizarInformaion(seccion[sec])
                print(f"Has seleccionado la seccion de{seccion[sec]}")
                df=pd.DataFrame(noticias)
                print(df)
                guardarInformacion(df,sec)
                palabrasClaves(noticias)
                salir=input(Fore.CYAN+"Deseas salir del scrip? (s/n)\n")
                if salir !='n':
                    print(Fore.WHITE+"Gracias por usar el script\n")
                    continuar=False

            else:
                print(Fore.RED + "Opcion no valida.")
            
        except  ValueError:
            print(Fore.RED + "Por favor, introduce un numero.")


#En este funcion verificamos si ya existe un archivo con nombre determinado y si no lo crea
def guardarInformacion(df,num):
    name=f"noticas_{nombres[num].replace(' ','_')}.csv"
    #Esta funcion que esta incluida en la libreria OS se encarga de revisar si existe un directorio con este nombre
    if os.path.exists(name):
        print("El archivo ya existe asi que sera abierto al final de el mismo para no eliminar los cambios\n")
        df.to_csv(name,mode='a',index=False,encoding='utf-8',header=False)
    else:
        print(f"El archivo no existe asi que sera creado{name}\n")
        df.to_csv(name,mode='a',index=False,encoding='utf-8')
    
    print("Generando reporte Excel...")
    exportarExcelBonito(df, nombres[num])
    Max_word=contarPalabra(df)
    print(f"{Max_word.head(15)}")
        
#Aqui resivimos que seccion quiere el usuario que sea scrapeada, cuando sera enviada a un ID de telegram y si la persona esta buscando que una palabra sea contenida en un titulo
def enviarInformacionRelevante():
    pass
#Esta funcion se encargara de consultar la base de datos de las noticias del dia,para encontrar unas KEYWORDS que el usuario desee revisar
def palabrasClaves(noticias):
    
    for noticia in noticias:
        titular_limpio = noticia['Titular'].lower()
    
        for palabra in keywords:
            palabra_limpia = palabra.lower().strip()
            if palabra_limpia in titular_limpio:
                print(f"ALERTA! Encontre'{palabra}' en: {noticia['Titular']}{noticia['Link']}")
                
                # Aqui llamaria a tu funcion de Telegram:
                # enviar_telegram(noticia['Titular'], noticia['Link'])
                
                # Rompemos el ciclo interno (break) para que no te avise 2 veces 
                break
    
#Esta es una de las funciones principales ya que  se encarga de hacer las peticiones http

def analizarInformaion(url):
    try:
        response=requests.get(url,headers=headers)
        response.raise_for_status()
        soup=BeautifulSoup(response.content,'html.parser')
        #Aqui ecaluamos todas la etiquetas 'div'que tenga esta clase
        contenedores=soup.find_all('div',class_='elementor-element elementor-element-114bf412 e-con-full e-flex e-con e-child')
        print(f"Se ha leido {len(contenedores)} contenedores")
        #Un array con todas las noticias recopiladas
        noticias=[]
        for i in contenedores:
            titulo=i.find('h2')
            fecha=i.find('time')
            if titulo:
                titulo_limpio=titulo.get_text(strip=True)
                link=i.find('a')
                if link:
                    fecha_texto = fecha.get_text(strip=True) if fecha else "Fecha no encontrada"
                    noticias.append({'Titular':titulo_limpio,'Fecha':fecha_texto,'Link':link.get('href')})
        if noticias:
            return noticias
        
       
        
    except HTTPError:
        print(f"Se ha presentado un error a la hora de conectar con la pagina {HTTPError}\n")
 

def contarPalabra(df):
    
    if df.empty or 'Titular' not in df.columns:
        print(Fore.RED + "No hay datos para contar palabras.")
        return pd.DataFrame() # Retorna tabla vacia
    
    # Convertir a string todos los titulares para evitar errores si hay datos vacios
    texto_completo = ' '.join(df['Titular'].astype(str).tolist())
    # Limpieza: quitar signos y pasar a minusculas
    texto_limpio = re.sub(r'[^\w\s]', '', texto_completo).lower()
    palabras = texto_limpio.split()
    palabras_filtradas = [
        palabra for palabra in palabras 
        if palabra not in STOPWORDS and len(palabra) > 2
    ]
    conteo_palabras = Counter(palabras_filtradas)
    
    df_conteo = pd.DataFrame(conteo_palabras.items(), columns=['Palabra', 'Cantidad'])
    df_ordenado = df_conteo.sort_values(by='Cantidad', ascending=False)
    
    return df_ordenado
    
if __name__ == "__main__":
    menu()