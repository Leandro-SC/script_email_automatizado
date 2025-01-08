import os
import smtplib
import pandas as pd
import tkinter as tk
import time
from tkinter import filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage  
from email import encoders
import imgkit  # pip install imgkit
import random

# Variables globales
ruta_salida_global = None

ruta_wkhtmltoimage = r""  # Reemplaza con la ruta correcta en tu sistema
config = imgkit.config(wkhtmltoimage=ruta_wkhtmltoimage)

def html_a_imagen(cuerpo_html, nombre_imagen):
    opciones = {
        'format': 'png',
        'encoding': 'UTF-8'
    }
    imgkit.from_string(cuerpo_html, nombre_imagen, options=opciones, config=config)

def enviar_correo(destinatario, nombre, archivo_adjuntar, correo_remitente, contraseña):
    try:
       
        mensaje = MIMEMultipart('related')
        mensaje['From'] = correo_remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = f'Asunto Personalizado  - {nombre}'

       #Puedes personalizar el cuerpo del correo con HTML y CSS inline 
        cuerpo_correo = f"""
        <html>
            <body style="font-family: Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0;">
                <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f4; padding: 20px;">
                    <tr>
                        <td align="center">
                            <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);">
                                <tr>
                                    <td align="center" style="background-color: #001f60; color: white; padding: 10px; border-radius: 10px 10px 0 0;">
                                        <h2 style="text-transform: uppercase; color: white; margin: 0;"></h2>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 20px; line-height: 1.6; color: #333333; text-align: justify;">
                                        <p text-align: center;>
                                        <br/>
                                        <br/>
                                        </p>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" style="padding: 20px;">
                                        <img src="cid:" alt="" style="max-width: 100%; height: auto;"/>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </body>
        </html>
        """
        parte_html = MIMEText(cuerpo_correo, 'html')
        mensaje.attach(parte_html)

        with open('./img/logo_i.png', 'rb') as img_file:
            mime_image = MIMEImage(img_file.read())
            mime_image.add_header('Content-ID', '<logo>')  
            mensaje.attach(mime_image)

        with open(archivo_adjuntar, 'rb') as adjunto:
            parte = MIMEBase('application', 'pdf')
            parte.set_payload(adjunto.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(archivo_adjuntar)}"')
            mensaje.attach(parte)

        print("Configurando el servidor SMTP...")
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        print("Iniciando sesión en el servidor SMTP...")
        servidor.login(correo_remitente, contraseña) 
        texto = mensaje.as_string()
        print(f"Enviando correo a {destinatario}...")
        servidor.sendmail(correo_remitente, destinatario, texto)
        servidor.quit()
        print("Correo enviado con éxito.")


        return True

    except smtplib.SMTPAuthenticationError:
        print(f'Error de autenticación. Verifica el correo y la contraseña de {correo_remitente}.')
        return False

    except Exception as e:
        print(f'Error al enviar correo a {destinatario}: {str(e)}')
        return False
    
def cargar_excel():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        try:
            df = pd.read_excel(archivo)
            return df, archivo
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
            return None, None
    return None, None

def enviar_correos():
    global ruta_salida_global
    df, archivo = cargar_excel()
    if df is not None:
        if 'nombre' in df.columns and 'correo' in df.columns and 'ruta' in df.columns and 'nombre_archivo' in df.columns:

            correo_remitente = entrada_remitente.get()
            contraseña = entrada_contraseña.get()

            if not correo_remitente or not contraseña:
                messagebox.showerror("Error", "Debe ingresar el correo remitente y la contraseña")
                return

            estados_envio = []

            for index, row in df.iterrows():
                nombre = row['nombre']
                correo= row['correo']
                ruta = row['ruta']
                nombre_archivo = row['nombre_archivo']
                
                archivo_personalizado = os.path.join(ruta, f"{nombre_archivo}.pdf")
                tiempo_random = random.randint(80, 100)
                time.sleep(tiempo_random)
                
                if not os.path.exists(archivo_personalizado):
                    messagebox.showerror("Error", f"No se encontró el archivo: {archivo_personalizado}")
                    estados_envio.append("No enviado")
                    continue

                enviado = enviar_correo(correo, nombre, archivo_personalizado, correo_remitente, contraseña)
                estados_envio.append("Enviado" if enviado else "No enviado")

            df['estado_envio'] = estados_envio
            ruta_salida_global = archivo.replace(".xlsx", "_resultado.xlsx")
            df.to_excel(ruta_salida_global, index=False)
            time.sleep(2)

            messagebox.showinfo("Éxito", f"Correos enviados correctamente. Resultado guardado en: {ruta_salida_global}")
            btn_descargar.config(state=tk.NORMAL)  

        else:
            messagebox.showerror("Error", "El archivo debe contener las columnas: nombre, correo, ruta, nombre_archivo")
    else:
        messagebox.showerror("Error", "No se pudo cargar el archivo")

def descargar_estado():
    if ruta_salida_global:
        filedialog.asksaveasfilename(initialfile=ruta_salida_global, defaultextension=".xlsx")
        messagebox.showinfo("Descarga", f"El archivo ha sido guardado en: {ruta_salida_global}")
    else:
        messagebox.showerror("Error", "No hay archivo disponible para descargar")



"""
Interfaz Gráfica del Programa
"""
root = tk.Tk()
root.title("Aplicación para Enviar Correos")
root.geometry("500x500")


tk.Label(root, text="Correo Remitente:").pack(pady=5)
entrada_remitente = tk.Entry(root, width=50)
entrada_remitente.pack(pady=5)

tk.Label(root, text="Contraseña:").pack(pady=5)
entrada_contraseña = tk.Entry(root, show="*", width=50)
entrada_contraseña.pack(pady=5)

btn_enviar = tk.Button(root, text="Cargar Excel y Enviar Correos", command=enviar_correos)
btn_enviar.pack(pady=20)

btn_descargar = tk.Button(root, text="Descargar Estado de Envíos", state=tk.DISABLED, command=descargar_estado)
btn_descargar.pack(pady=10)

root.mainloop()


