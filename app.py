import os
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import calendar
import locale
from tkinter.ttk import Progressbar

# Establecer configuración regional en español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')


class SeparadorLiquidaciones:
    def __init__(self):
        self.ventana = tk.Tk()
        self.ventana.title(
            "Separador de Liquidaciones PDF y Envío de Correos Electrónicos")
        self.ventana.geometry("500x200")

        # Etiqueta y botón para seleccionar el archivo PDF y separar las páginas
        etiqueta = tk.Label(self.ventana, text="Seleccionar archivo PDF:")
        etiqueta.pack()

        boton_separar_pdf = tk.Button(
            self.ventana, text="Separar PDF", command=self.separar_pdf)
        boton_separar_pdf.pack(pady=10)

        # Botón para enviar los correos electrónicos
        boton_enviar_correos = tk.Button(
            self.ventana, text="Enviar Correos", command=self.enviar_correos)
        boton_enviar_correos.pack(pady=10)

        # Mostrar instrucciones
        self.mostrar_instrucciones()

    def procesar_texto_para_extraer_nombre(self, texto):
        # Patrón de búsqueda del nombre
        patron_nombre = r"Nombre:\s+(.+)"

        # Buscar el nombre utilizando expresiones regulares
        coincidencia = re.search(patron_nombre, texto)

        if coincidencia:
            return coincidencia.group(1)
        else:
            return None

    def separar_pdf(self):
        # Obtener la ruta del archivo PDF
        ruta_pdf = filedialog.askopenfilename(
            filetypes=[("Archivos PDF", "*.pdf")])

        if not ruta_pdf:
            return

        try:
            # Obtener el mes correspondiente
            mes_correspondiente = self.obtener_mes_correspondiente()

            if not mes_correspondiente:
                return

            # Obtener el nombre del mes y el año
            nombre_mes_correspondiente = calendar.month_name[mes_correspondiente]
            nombre_mes_correspondiente = nombre_mes_correspondiente.capitalize()
            mes_correspondiente_completo = f"{nombre_mes_correspondiente.capitalize()} {datetime.now().year}"

            # Crear una carpeta para el mes correspondiente (si no existe)
            carpeta_mes_correspondiente = f'Liquidaciones/{mes_correspondiente}.- {mes_correspondiente_completo}'
            os.makedirs(carpeta_mes_correspondiente, exist_ok=True)

            # Instanciar el lector de PDF
            lector_pdf = PdfReader(ruta_pdf)

            # Procesar cada página del PDF
            for num_pagina, pagina in enumerate(lector_pdf.pages):
                # Extraer texto de la página
                contenido = pagina.extract_text()

                # Buscar el nombre del trabajador en el texto utilizando técnicas de procesamiento de texto
                nombre = self.procesar_texto_para_extraer_nombre(contenido)

                # Si se encontró el nombre, guardar la página en el PDF correspondiente al trabajador
                if nombre:
                    # Generar el nombre del archivo para el trabajador actual
                    nombre_archivo = f'{nombre}.pdf'

                    # Ruta completa del PDF del trabajador
                    ruta_trabajador_pdf = f'{carpeta_mes_correspondiente}/{nombre_archivo}'

                    # Crear un nuevo PDF solo con la página actual
                    escritor_pdf = PdfWriter()
                    escritor_pdf.add_page(pagina)

                    # Guardar el PDF del trabajador
                    with open(ruta_trabajador_pdf, 'wb') as archivo_salida:
                        escritor_pdf.write(archivo_salida)

            messagebox.showinfo("Proceso completado",
                                "Las páginas se han separado correctamente.")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Ocurrió un error durante el procesamiento del PDF: {str(e)}")

    def enviar_correos(self):
        # Obtener el mes correspondiente
        mes_correspondiente = self.obtener_mes_correspondiente()

        if not mes_correspondiente:
            return

        # Obtener el nombre del mes y el año
        nombre_mes_correspondiente = calendar.month_name[mes_correspondiente]
        nombre_mes_correspondiente = nombre_mes_correspondiente.capitalize()
        mes_correspondiente_completo = f"{nombre_mes_correspondiente.capitalize()} {datetime.now().year}"

        # Obtener la carpeta correspondiente al mes correspondiente
        carpeta_mes_correspondiente = f'Liquidaciones/{mes_correspondiente}.- {mes_correspondiente_completo}'

        # Obtener la lista de archivos PDF en la carpeta
        archivos_pdf = [archivo for archivo in os.listdir(
            carpeta_mes_correspondiente) if archivo.endswith(".pdf")]

        if not archivos_pdf:
            messagebox.showwarning(
                "Advertencia", "No hay archivos PDF para enviar.")
            return

        def actualizar_progreso(porcentaje):
            progreso["value"] = porcentaje
            ventana_progreso.update_idletasks()

        try:
            # Configuración del correo electrónico
            remitente = "Corre_de_Gmail" 
            contraseña = "Contraseña_Google" ### Contraseña para aplicaciones de google, no es la contraseña normal.

            # Crear ventana de progreso
            ventana_progreso = tk.Toplevel(self.ventana)
            ventana_progreso.title("Progreso")
            ventana_progreso.geometry("300x100")
            # Establecer la ventana como la de primer plano
            ventana_progreso.attributes("-topmost", True)

            # Etiqueta y barra de progreso
            etiqueta_progreso = tk.Label(
                ventana_progreso, text="Enviando correos electrónicos...")
            etiqueta_progreso.pack(pady=10)

            progreso = Progressbar(
                ventana_progreso, length=200, mode='determinate')
            progreso.pack()

            for indice, archivo_pdf in enumerate(archivos_pdf, start=1):
                # Ruta completa del PDF del trabajador
                ruta_trabajador_pdf = f'{carpeta_mes_correspondiente}/{archivo_pdf}'

                # Obtener el nombre del trabajador a partir del nombre del archivo
                nombre_trabajador = os.path.splitext(archivo_pdf)[0]

                diccionario_correos = {
                    "Nombre": "Correo",
                  

                }

                # Obtener el correo electrónico del trabajador
                correo_trabajador = diccionario_correos.get(nombre_trabajador)
                correo_copias = [""]
                destinatarios = [correo_trabajador] + list(correo_copias)

                if correo_trabajador:
                    # Configurar el correo electrónico
                    mensaje = MIMEMultipart()
                    mensaje["From"] = remitente
                    mensaje["To"] = ", ".join(destinatarios)
                    mensaje["Subject"] = f"Liquidación de Sueldo {nombre_mes_correspondiente} {datetime.now().year}"


                    cuerpo = f"""
<html>
<body>
    <p>Estimado(a) {nombre_trabajador},</p>
    <p>Adjunto liquidación correspondiente al mes de {nombre_mes_correspondiente} {datetime.now().year}.</p>
    <p>Saludos cordiales,</p>
    <div dir="ltr"><span style="font-size:12.8px">Atte.</span><div style="font-size:12.8px"><div dir="ltr"><div dir="ltr"><div dir="ltr"><table style="font-size:medium;font-family:Verdana,sans-serif;color:rgb(10,85,146)"><tbody><tr><td style="font-weight:bold;font-size:15px;color:rgb(255,68,56)"><table style="font-weight:400;font-size:medium;color:rgb(10,85,146)"><tbody><tr><td style="font-weight:bold;font-size:15px;color:rgb(255,68,56)">Carlos M. Álvarez Salas</td></tr></tbody></table></td></tr><tr><td style="font-size:13px">Jefe de Servicio ITO OCA Global<br>División Industria &amp; NDT<br></td></tr><tr><td height="5"><img src="https://ci3.googleusercontent.com/proxy/pFMuPOIH5NdX9UigyKb-PyafYkc4O_U7ULfLY8kZAujKSR_mIqHfIWpGdxweA_MQmyOAVQXGMkB2KxUY7mZFmbBQ=s0-d-e1-ft#http://ocaglobal.com/firmamail/img/logo_oca.png" alt="logo_oca" width="238" height="65" style="color:rgb(17,85,204)" class="CToWUd" data-bit="iit"><br></td></tr><tr><td><br></td></tr><tr><td height="5"></td></tr><tr><td style="font-size:12px"><img src="https://ci5.googleusercontent.com/proxy/Pj-CEKOGtvRJBJvDPJebHHrr-rVhAAVcvKJk2YIdkzhp7wrYfsqokck4gLLKuKs7ZhjiJxKe_RIqakgniW32eF_VY8f65yFb=s0-d-e1-ft#https://ocaglobal.cl/firmamail/img/icono_telefono.png" alt="icono_telefono" width="25" height="15" class="CToWUd" data-bit="iit">&nbsp;<font color="#0000ff">+(56) 9 4869 4963</font></td></tr><tr><td style="font-size:12px"><img src="https://ci3.googleusercontent.com/proxy/xg3h1FjNWbQxz-2xn1mKZh6W2bZR9Lgo5jphjtUXCi8wIMh1yTARwituNpN5pTWFgp9Ewkbd4MIz__fIF5askERRzc2U=s0-d-e1-ft#https://ocaglobal.cl/firmamail/img/icono_email.png" alt="icono_email" width="25" height="9" class="CToWUd" data-bit="iit">&nbsp;&nbsp;<font color="#0000ff">carlos.alvarez<a>@ocaglobal.com</a></font></td></tr><tr><td style="font-size:12px"><img src="https://ci6.googleusercontent.com/proxy/t7R1IK4Pn7r0u6PGSS75ndVU6rZ4hK7Z3kNJyFsaHvifDZczRvCt9-RRfNqo84B7zyoUWT2EhtWVpR8SDhbvHZLOgQ=s0-d-e1-ft#https://ocaglobal.cl/firmamail/img/icono_web.png" alt="icono_web" width="25" height="13" class="CToWUd" data-bit="iit">&nbsp;&nbsp;<a href="http://www.ocaglobal.com/" style="color:rgb(10,85,146)" target="_blank" data-saferedirecturl="https://www.google.com/url?q=http://www.ocaglobal.com/&amp;source=gmail&amp;ust=1685924718788000&amp;usg=AOvVaw1iVoQ7fHFPzMr1Vrd0fSkw">www.ocaglobal.com</a></td></tr><tr><td height="6"><a href="https://maps.google.com/?q=Av.+Pedro+de+Valdivia+291+Piso+9+Providencia,+Santiago+de+Chile&amp;entry=gmail&amp;source=g" style="font-size:12px;color:rgb(17,85,204)" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://maps.google.com/?q%3DAv.%2BPedro%2Bde%2BValdivia%2B291%2BPiso%2B9%2BProvidencia,%2BSantiago%2Bde%2BChile%26entry%3Dgmail%26source%3Dg&amp;source=gmail&amp;ust=1685924718789000&amp;usg=AOvVaw1Y2-jl32BgWL8MqDkrlKrP">Av. Pedro de Valdivia 291 Piso 9</a><br style="font-size:12px"><a href="https://maps.google.com/?q=Av.+Pedro+de+Valdivia+291+Piso+9+Providencia,+Santiago+de+Chile&amp;entry=gmail&amp;source=g" style="font-size:12px;color:rgb(17,85,204)" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://maps.google.com/?q%3DAv.%2BPedro%2Bde%2BValdivia%2B291%2BPiso%2B9%2BProvidencia,%2BSantiago%2Bde%2BChile%26entry%3Dgmail%26source%3Dg&amp;source=gmail&amp;ust=1685924718789000&amp;usg=AOvVaw1Y2-jl32BgWL8MqDkrlKrP">Providencia, Santiago de Chile</a><br></td></tr><tr><td height="3"></td></tr><tr><td><a href="https://es.linkedin.com/company/oca-global" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://es.linkedin.com/company/oca-global&amp;source=gmail&amp;ust=1685924718789000&amp;usg=AOvVaw3P64eX4N_3dIaDR9R4WXmf"><img src="https://ci6.googleusercontent.com/proxy/H3wLn9A9vLiB4KmSi2Qb51boH8ImvUMasX1l6Nj03ttukVcplXNBe7djdGdS-ZYUDRH2lohoJEhlXh4SP3hrfn31yhEq9eHq=s0-d-e1-ft#https://ocaglobal.cl/firmamail/img/icono_linkedin.png" alt="icono_linkedin" width="29" height="29" class="CToWUd" data-bit="iit"></a></td></tr><tr><td height="10"></td></tr><tr><td><a href="https://odisdkp.firebaseapp.com/#/home" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://odisdkp.firebaseapp.com/%23/home&amp;source=gmail&amp;ust=1685924718789000&amp;usg=AOvVaw0iA5oHGTZCAEeb679uI0pS"><img src="https://ci4.googleusercontent.com/proxy/V_lyOrYwQlxeAb3rwGPMWKPkaC3dV7fuoP0Pcx7XUp9aOjOHkADxMI_M2dPcExq5vv-E_kQialTIkI__vWIKaYxRuUNCZJnpkl-D=s0-d-e1-ft#https://ocaglobal.cl/firmamail/img/banner_odis_chile.jpg" alt="Banner OCA Global" width="600" height="200" class="CToWUd" data-bit="iit"></a></td></tr></tbody></table></div></div></div></div></div></body>
</html>
"""

                    mensaje.attach(MIMEText(cuerpo, "html"))

                    with open(ruta_trabajador_pdf, "rb") as adjunto:
                        parte_adjunto = MIMEApplication(
                            adjunto.read(), Name=os.path.basename(ruta_trabajador_pdf))
                        parte_adjunto[
                            "Content-Disposition"] = f'attachment; filename="{os.path.basename(ruta_trabajador_pdf)}"'
                        mensaje.attach(parte_adjunto)

                    # Enviar el correo electrónico
                    servidor_smtp = smtplib.SMTP("smtp.gmail.com", 587)
                    servidor_smtp.starttls()
                    servidor_smtp.login(remitente, contraseña)
                    servidor_smtp.sendmail(
                        remitente, destinatarios, mensaje.as_string())
                    servidor_smtp.quit()

                # Actualizar la barra de progreso
                porcentaje = (indice / len(archivos_pdf)) * 100
                actualizar_progreso(porcentaje)

            messagebox.showinfo(
                "Envío de Correos", "Los correos electrónicos se han enviado correctamente.")
        except Exception as e:
            messagebox.showerror("Error",
                                 f"Ocurrió un error durante el envío de correos electrónicos: {str(e)}")

    def mostrar_instrucciones(self):
        mensaje = """
        Bienvenido al Separador de Liquidaciones PDF y Envío de Correos Electrónicos.

        Antes de comenzar, asegúrese de tener un archivo PDF válido para separar las páginas.

        Pasos a seguir:
        1. Haga clic en el botón 'Seleccionar archivo PDF' para elegir el archivo PDF.
        2. Una vez seleccionado el archivo, se le pedirá que ingrese el número del mes al que corresponden las liquidaciones (1-12).
        3. Después de ingresar el mes, haga clic en el botón 'Separar PDF' para separar las páginas del PDF en archivos individuales para ese mes.
        4. Luego de separar las páginas, puede hacer clic en el botón 'Enviar Correos' para enviar los correos electrónicos con los archivos adjuntos del mes seleccionado.
        5. Si necesita agregar un nuevo trabajador, por favor, notifíquelo a Diego Bravo para que pueda actualizar el envío de correos electrónicos.

        ¡Recuerde ingresar las direcciones de correo electrónico correctas en el diccionario 'diccionario_correos' antes de enviar los correos electrónicos!
        ¡El envío de correos solo funcionará si existen liquidaciones separadas!

        Presione OK para comenzar.
        """

        messagebox.showinfo("Instrucciones", mensaje)

    def obtener_mes_correspondiente(self):
        # Obtener el mes correspondiente ingresado por el usuario
        mes_correspondiente = tk.simpledialog.askinteger("Mes correspondiente",
                                                         "Ingrese el número del mes correspondiente (1-12):",
                                                         parent=self.ventana,
                                                         minvalue=1, maxvalue=12)

        return mes_correspondiente

    def iniciar_aplicacion(self):
        # Ejecutar el bucle principal de la ventana
        self.ventana.mainloop()


# Crear instancia de la clase y ejecutar la aplicación
app = SeparadorLiquidaciones()
app.iniciar_aplicacion()
