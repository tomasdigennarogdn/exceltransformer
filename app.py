import streamlit as st
import pandas as pd
from gtts import gTTS
from time import sleep
from datetime import datetime
import os
import pyglet
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

def mensaje_final(texto = None):
    if not texto:
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        texto = f'proceso terminado a las {current_time}'
            
    filename = 'mensajetemp.mp3'
    if os.path.exists(filename):
        os.remove(filename)

    tts = gTTS(text= texto, lang='es')
    tts.save(filename)
    music = pyglet.media.load(filename, streaming=False)
    
    music.play()
    
    sleep(music.duration) 
    os.remove(filename)

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def mensaje_descarga():
	if muteado == False:
		mensaje_final("Atención!. Archivo descargado. Hora de agarrar la pala")

##########################################
##########################################

st.markdown("# transformador de columna de excel.")

st.markdown("### Lo que te proponía, hasta que podamos hablar con simplyroute es algo de este estilo")

st.write("Primero, elijamos algunas opciones:")

respuesta_radio = st.radio("Cambiamos la columna Id de referencia?", ("Si","No"))
if respuesta_radio == "Si":
	col = "Id de referencia"
if respuesta_radio == "No":
	col = st.text_input("Por favor, poné que columna querés, respetá los espacios y mayúsculas").strip()
	if col:
		st.write(f"ok, trabajando sobre la columna {col}")

respuesta_radio2 = st.radio("Por qué valor vamos a reemplazar los valores no numéricos?", ("default: -10","Otro"))
if respuesta_radio2 == "default: -10":
	reemplazo = -10
if respuesta_radio2 == "Otro":
	reemplazo = st.text_input("Por favor, poné qué valor querés (numérico)")
	reemplazo = int(reemplazo.strip())

respuesta_radio3 = st.radio("Querés que se notifique por audio cuando terminamos los procesos?", ("No","Si"))
if respuesta_radio3 == "Si":
	muteado = True
if respuesta_radio3 == "No":
	muteado = False

data = st.file_uploader("Por favor cargue el excel SOLAMENTE CON LA SOLAPA VISITAS. " , help = "lol", type = "xlsx")
if data:
	data = pd.read_excel(data)
	st.write("dataframe cargado satisfactoriamente")

	st.write("Vamos a cambiar el id de referencia")

	if not col in data.columns:
		st.error("Error! no se encontró la columna. recargar la página o consultar con tomi.")

	#extraigo los campos excepcionales y los inserto 
	nueva_col = data[col].apply(lambda x : "-" if type(x) in [int, float] else x )
	
	#extraigo donde está la columna a analizar
	pos = list(data).index(col)

	#inserto a la derecha la nuev columna
	data.insert(pos+1, "status id", nueva_col)

	#cambio la columna original insertando el reemplazo, sacando los valores raros 
	data[col] = data[col].apply(lambda x : x if type(x) in [int, float] else reemplazo )

	st.success("Archivo transformado satisfactoriamente!, a descargarlo!")
	if muteado == False:
		mensaje_final("Proceso terminado, disculpe la demora. Descargue el archivo por favor.")

	


	df_xlsx = to_excel(data)
	st.download_button(label='📥 click acá para descargar el archivo!',
                       data=df_xlsx ,
                       file_name= 'Visitas_transformado.xlsx',
                       on_click = mensaje_descarga)
	