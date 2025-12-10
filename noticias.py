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
#---------------------URLS----------------------------#
seccion={
    1:'https://www.el-carabobeno.com/secciones/noticias/nacional/',
    2:'https://www.el-carabobeno.com/secciones/noticias/internacional/',
    3:'https://www.el-carabobeno.com/secciones/deportes/',
    4:'https://www.el-carabobeno.com/secciones/noticias/valencia/',
    5:'https://www.el-carabobeno.com/secciones/universidad/',
}
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
            else:
                print(Fore.RED + "Opcion no valida.")
            
        except  ValueError:
            print(Fore.RED + "Por favor, introduce un numero.")


#En este funcion verificamos si ya existe un archivo con nombre determinado y si no lo crea
def guardarInformacion(df,num):
    name=f'noticas_{nombres[num].replace(' ','_')}.csv'
    #Esta funcion que esta incluida en la libreria OS se encarga de revisar si existe un directorio con este nombre
    if os.path.exists(name):
        print("El archivo ya existe asi que sera abierto al final de el mismo para no eliminar los cambios\n")
        df.to_csv(name,mode='a',index=False,encoding='utf-8',header=False)
    else:
        print(f"El archivo no existe asi que sera creado{name}\n")
        df.to_csv(name,mode='a',index=False,encoding='utf-8')
        
#Aqui resivimos que seccion quiere el usuario que sea scrapeada, cuando sera enviada a un ID de telegram y si la persona esta buscando que una palabra sea contenida en un titulo
def enviarInformacionRelevante():
    pass
#Esta es la funcion principal ya que se encarga de hacer las peticiones http
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
 
 
    
if __name__ == "__main__":
    menu()