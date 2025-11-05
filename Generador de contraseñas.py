import random, pandas, smtplib, os
from email.message import EmailMessage

# FUNCION para detectar que columna contiene los emails

def detectar_columna(planilla_de_excel):

    columnas = []

    for c in planilla_de_excel.columns:
        columnas.append(c.lower())

    for columna_correcta in ['email', 'correo', 'mail', 'correos', 'emails', 'mails']:
        if columna_correcta in columnas:
            return planilla_de_excel.columns[columnas.index(columna_correcta)]
    return planilla_de_excel.columns[0]


# FUNCION para generar contraseñas aleatorias

def generar_contrasenha():
    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    symbols = ['!', '#', '$', '%', '&', '(', ')', '*', '+']

    cantidad_de_letras = int(random.uniform(2, 5))
    cantidad_de_numeros = int(random.uniform(2, 5))
    cantidad_de_simbolos = int(random.uniform(2, 5))

    caracteres_para_lista = []

    for letra in range(0, cantidad_de_letras):
        caracteres_para_lista.append(random.choice(letters))
    for numero in range(0, cantidad_de_numeros):
        caracteres_para_lista.append(random.choice(numbers))
    for simbolo in range(0, cantidad_de_simbolos):
        caracteres_para_lista.append(random.choice(symbols))

    lista_final = []

    for i in range(len(caracteres_para_lista) - 1, -1, -1):
        letra_guardada = random.choice(caracteres_para_lista)
        lista_final.append(letra_guardada)
        if letra_guardada in caracteres_para_lista:
            caracteres_para_lista.remove(letra_guardada)

    return "".join(lista_final)

# Se le pide al usuario que meta el nombre de su archivo
nombre = input("Ingrese el nombre del archivo excel (sin .xlsx): ").strip()

# Si el usuario ya ingreso el .xlsx entonces se omite esta parte, si no ingreso, se le añade al final el .xlsx
if nombre[:-5] != ".xlsx":
    nombre += ".xlsx"

directorio_del_programa = os.path.dirname(os.path.abspath(__file__)) #Este es el directorio del codigo, en el cual se debe encontrar tambien el excel para que funcione
archivo = os.path.join(directorio_del_programa, nombre)
planilla = pandas.read_excel(archivo, engine='openpyxl')

titulo_de_columna_de_correos = detectar_columna(planilla)


# Creamos la nueva columna para las contraseñas

correos_faltantes = []

if "Password" not in planilla.columns:
    planilla["Password"] = None

for i in range(len(planilla)): 
    if pandas.isnull(planilla.loc[i, "Password"]) or str(planilla.loc[i, "Password"]).strip() == "":
        planilla.loc[i, "Password"] = generar_contrasenha()
        correos_faltantes.append(str(planilla.loc[i, titulo_de_columna_de_correos]).strip())
    else:
        print(f"{str(planilla.loc[i, titulo_de_columna_de_correos])} ya tiene una contrasena")
        continue
 
planilla.to_excel(archivo, index=False, engine='openpyxl')


# Envio de correos

EMAIL_REMITENTE = "nicomingouptp@gmail.com"
CLAVE_APP = "acubadbevcypdpmk"

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL_REMITENTE, CLAVE_APP)
    for _, fila in planilla.iterrows():

        correo = str(fila[titulo_de_columna_de_correos]).strip()
        contrasena = fila["Password"]
         
        if "@" not in correo:
            continue
        elif correo in correos_faltantes:
            msg = EmailMessage()
            msg["Subject"] = "Tu nueva contraseña"
            msg["From"] = EMAIL_REMITENTE
            msg["To"] = correo
            msg.set_content(f"Hola\n\nTus nueva contraseña es:\n{contrasena}\n\nSaludos.")

            smtp.send_message(msg)
            print(f"Enviado a {correo}")
        else:
            continue