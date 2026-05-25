from flask import Flask, render_template, request
import os
from datetime import datetime
import openpyxl  # Librería que genera el Excel
from openpyxl import Workbook
from preguntas import lista_preguntas  # Importamos el archivo de la base de datos

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def inicio():
    if request.method == 'POST':
        aciertos = 0
        respuestas_usuario = {}
        
        for p in lista_preguntas:
            letra_elegida = request.form.get(str(p['id']), "Sin responder")
            respuestas_usuario[p['id']] = letra_elegida
            if letra_elegida == p['respuesta_correcta']:
                aciertos += 1
                
        calificacion = round((aciertos / len(lista_preguntas)) * 10, 1)
        
        # --- AQUÍ GENERAMOS EL ARCHIVO EXCEL (.XLSX) 100% NATIVO ---
        archivo_excel = 'Resultados_Estudiantes.xlsx'
        existe = os.path.isfile(archivo_excel)
        
        # Si el Excel ya existe, lo abrimos; si no, creamos uno nuevo
        if existe:
            wb = openpyxl.load_workbook(archivo_excel)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Calificaciones"
            # Creamos los encabezados de la tabla
            encabezados = ['Fecha', 'Aciertos', 'Calificación'] + [f'Preg {p["id"]}' for p in lista_preguntas]
            ws.append(encabezados)
        
        # Añadimos la nueva fila con los resultados de la alumna
        fila = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), aciertos, calificacion]
        for p in lista_preguntas:
            fila.append(respuestas_usuario[p['id']])
            
        ws.append(fila)
        wb.save(archivo_excel)  # Guardamos el archivo .xlsx
        
        # Pantalla final de éxito
        resultado_html = f"""
        <body style='background-color:#0b0c10; color:#c5c6c7; font-family:sans-serif; text-align:center; padding-top:100px;'>
            <h1 style='color:#66fcf1;'>¡EVALUACIÓN GUARDADA CON ÉXITO!</h1>
            <h2 style='color:#ffffff;'>Aciertos: {aciertos} de {len(lista_preguntas)}</h2>
            <h2 style='color:#45a29e;'>Calificación: {calificacion}</h2>
            <p>Tus respuestas se han exportado al Excel del profesor.</p>
        </body>
        """
        return resultado_html

    return render_template('examen.html', preguntas=lista_preguntas)

if __name__ == '__main__':
    app.run(debug=True, port=5050)
import os

if __name__ == '__main__':
    # Render asigna un puerto automáticamente en la variable de entorno 'PORT'
    port = int(os.environ.get('PORT', 5050))
    # Es vital usar host='0.0.0.0' para que Render pueda conectar
    app.run(host='0.0.0.0', port=port)