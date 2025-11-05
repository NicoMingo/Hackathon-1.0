import random
import pandas as pd
import smtplib
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email.message import EmailMessage

# ------------------------------
# Funciones de procesamiento
# ------------------------------

def detectar_columna(planilla_de_excel):
    columnas = [c.lower() for c in planilla_de_excel.columns]
    for columna_correcta in ['email', 'correo', 'mail', 'correos', 'emails', 'mails']:
        if columna_correcta in columnas:
            return planilla_de_excel.columns[columnas.index(columna_correcta)]
    return planilla_de_excel.columns[0]

def generar_contrasenha():
    letters = list("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
    numbers = list("0123456789")
    symbols = list("!#$%&()*+")

    cantidad_de_letras = int(random.uniform(2, 5))
    cantidad_de_numeros = int(random.uniform(2, 5))
    cantidad_de_simbolos = int(random.uniform(2, 5))

    caracteres_para_lista = (
        [random.choice(letters) for _ in range(cantidad_de_letras)] +
        [random.choice(numbers) for _ in range(cantidad_de_numeros)] +
        [random.choice(symbols) for _ in range(cantidad_de_simbolos)]
    )

    lista_final = []
    while caracteres_para_lista:
        letra_guardada = random.choice(caracteres_para_lista)
        lista_final.append(letra_guardada)
        caracteres_para_lista.remove(letra_guardada)

    return "".join(lista_final)

def procesar_excel(ruta_excel):
    print(f"[INFO] Procesando archivo: {ruta_excel}")
    planilla = pd.read_excel(ruta_excel, engine='openpyxl')
    titulo_de_columna_de_correos = detectar_columna(planilla)

    registro_contrasenas = {}  # correo: contraseña
    correos_a_enviar = []      # solo los correos nuevos que requieren mail

    if "Password" not in planilla.columns:
        planilla["Password"] = None

    for i in range(len(planilla)):
        correo_actual = str(planilla.loc[i, titulo_de_columna_de_correos]).strip()

        if "@" not in correo_actual:
            print(f"[IGNORADO] Fila {i+1}: correo inválido → {correo_actual}")
            continue

        if correo_actual in registro_contrasenas:
            planilla.loc[i, "Password"] = registro_contrasenas[correo_actual]
            print(f"[REUTILIZADO] Fila {i+1}: {correo_actual} ya tiene contraseña → misma usada")
            continue

        if pd.notnull(planilla.loc[i, "Password"]) and str(planilla.loc[i, "Password"]).strip() != "":
            registro_contrasenas[correo_actual] = str(planilla.loc[i, "Password"]).strip()
            print(f"[EXISTENTE] Fila {i+1}: {correo_actual} ya tenía contraseña → se respeta")
            continue

        nueva_pass = generar_contrasenha()
        planilla.loc[i, "Password"] = nueva_pass
        registro_contrasenas[correo_actual] = nueva_pass
        correos_a_enviar.append(correo_actual)
        print(f"[NUEVO] Fila {i+1}: {correo_actual} → contraseña generada")

    planilla.to_excel(ruta_excel, index=False, engine='openpyxl')
    print("[INFO] Contraseñas actualizadas y Excel guardado.\n")

    # ------------------------------
    # Envío de correos
    # ------------------------------
    EMAIL_REMITENTE = "nicomingouptp@gmail.com"
    CLAVE_APP = "acubadbevcypdpmk"

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_REMITENTE, CLAVE_APP)
        for correo in correos_a_enviar:
            contrasena = registro_contrasenas[correo]
            msg = EmailMessage()
            msg["Subject"] = "Tu nueva contraseña"
            msg["From"] = EMAIL_REMITENTE
            msg["To"] = correo
            msg.set_content(f"Hola,\n\nTu nueva contraseña es:\n{contrasena}\n\nSaludos.")

            try:
                smtp.send_message(msg)
                print(f"[ENVIADO] {correo}")
            except Exception as e:
                print(f"[ERROR] No se pudo enviar a {correo} → {e}")

    print("[INFO] Proceso completado.")

# ------------------------------
# Funciones de Tkinter
# ------------------------------

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if ruta:
        archivo_entry.delete(0, tk.END)
        archivo_entry.insert(0, ruta)

def enviar_contraseñas():
    ruta = archivo_entry.get().strip()
    if not ruta or not os.path.isfile(ruta):
        messagebox.showerror("Error", "Selecciona un archivo válido")
        return

    progreso["value"] = 0
    root.update_idletasks()

    try:
        procesar_excel(ruta)
        progreso["value"] = 100
        messagebox.showinfo("Listo", "Las contraseñas fueron enviadas correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        progreso["value"] = 0

# ------------------------------
# GUI Tkinter
# ------------------------------

root = tk.Tk()
root.title("Launcher - Enviar Contraseñas")
root.geometry("500x200")
root.eval('tk::PlaceWindow . center')

archivo_label = tk.Label(root, text="Archivo Excel:", font=("Arial", 12))
archivo_label.pack(pady=10)

archivo_entry = tk.Entry(root, width=50, font=("Arial", 12))
archivo_entry.pack(pady=5)

seleccionar_btn = tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo, font=("Arial", 12))
seleccionar_btn.pack(pady=5)

enviar_btn = tk.Button(root, text="Enviar Contraseñas", command=enviar_contraseñas, font=("Arial", 12), bg="#4CAF50", fg="white")
enviar_btn.pack(pady=10)

progreso = ttk.Progressbar(root, length=400, mode='determinate')
progreso.pack(pady=5)

root.mainloop()


