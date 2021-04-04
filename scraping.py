import requests
import pandas as pd
import xlsxwriter
from bs4 import BeautifulSoup
from datetime import date

fecha = date.today().strftime('%d/%m/%Y')
baseURL = 'https://www.dof.gob.mx/indicadores_detalle.php?cod_tipo_indicador=158&dfecha=07/01/2021&hfecha=07/01/2021'
baseURL = 'https://www.dof.gob.mx/indicadores_detalle.php?cod_tipo_indicador=158&dfecha='+fecha+'&hfecha='+fecha
page = requests.get(baseURL)
soup = BeautifulSoup(page.content, 'html.parser')
tipoCambioElements = soup.find_all('td', class_= 'txt')
if len(tipoCambioElements) == 4:
    tipoCambio = tipoCambioElements[3].text
else:
    tipoCambio = 'N.A'
# ESCRIBIR EL TIPO DE CAMBIO EN EL LIBRO DE EXCEL
mesEspanol = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
dia = int(date.today().strftime('%d'))
numeroMes = int(date.today().strftime('%m'))
nombreMes = mesEspanol[numeroMes-1]
df = pd.read_excel('D:\Development\Python\Tipo de cambio\plantilla_tc.xlsx',0, index_col=0)
df.loc[[dia],[nombreMes]] = tipoCambio

try:
    #CREAMOS UN NUEVO ARCHIVO DE EXCEL
    writer = pd.ExcelWriter('D:\Development\Python\Tipo de cambio\Tipo de cambio 2021.xlsx', engine='xlsxwriter')
    df.to_excel(writer,sheet_name='2021')
    libro = writer.book
    formatCell = libro.add_format({'bg_color':'#DAF7A6'})
    hojaExcel = writer.sheets['2021']
    hojaExcel.conditional_format('A2:A32', {'type': '3_color_scale'})
    hojaExcel.conditional_format('B1:M1', {'type': 'no_blanks', 'format': formatCell})
    #PINTAR LA CELDA VACIA
    hojaExcel.conditional_format('A1', {'type': 'blanks', 'format': formatCell})

    writer.save()
    df.to_excel('D:\Development\Python\Tipo de cambio\plantilla_tc.xlsx')
except IOError:
    print('Cerrar los archivos')
