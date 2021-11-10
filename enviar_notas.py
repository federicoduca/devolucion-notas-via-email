import smtplib
from email.message import EmailMessage
import os
import pandas as pd
from string import Template

# Variables de configuración
# Datos Gmail
GMAIL = dict()
GMAIL["ADDRESS"] = os.environ.get('GMAIL_USER')
GMAIL["PASSWORD"] = os.environ.get('GMAIL_PASS')
GMAIL["SERVER"] = 'smtp.gmail.com'
GMAIL["PORT"] = 587
# Datos OFFICE
OFFICE = dict()
OFFICE["ADDRESS"] = os.environ.get('OFFICE_USER')
OFFICE["PASSWORD"] = os.environ.get('OFFICE_PASS')
OFFICE["SERVER"] = 'smtp.office365.com'
OFFICE["PORT"] = 587
# Datos mail profesor:
MAIL_PROFESOR = "profesor@mail.edu"
# Contenido de notas
PATH_NOTAS = 'NOTAS - Examen.csv'
CAMPOS_RELEVANTES = ['Legajo', 'Nombre', 'Apellido', 'Dirección de correo', 'Entregó', 'Nota ej1 (30 puntos)',
           'Observaciones ej1', 'D&C', 'Complejidad', 'Nota ej2 (40 puntos)', 'Observaciones ej2', 'Nota ej3 (30 puntos)', 'Observaciones ej3', 'Nota Final', 'Comentarios']


if __name__ == '__main__':

  # cambiamos directorio de trabajo
  abspath = os.path.abspath(__file__)
  os.chdir(os.path.dirname(abspath))

  print("Cargando notas... ", end="")
  df_notas = pd.read_csv(PATH_NOTAS, encoding="UTF-8",
               usecols=CAMPOS_RELEVANTES)
  print("Listo!")

  # Nos quedamos solo con el primer nombre y apellido
  df_notas["Nombre Completo"] = df_notas["Nombre"] + \
    " " + df_notas["Apellido"]
  df_notas["Nombre"] = df_notas["Nombre"].apply(lambda x: x.split(' ')[0])
  df_notas["Apellido"] = df_notas["Apellido"].apply(
    lambda x: x.split(' ')[0])
  df_notas["Aprobado"] = df_notas["Nota Final"].apply(
    lambda x: "Aprobado" if x >= 60 else "No Aprobado")

  # Conectando al servidor:
  # servidor = "OFFICE"
  servidor = "GMAIL"
  MAIL = GMAIL if servidor == "GMAIL" else OFFICE
  print(f"Conectando al servidor: {servidor}")

  with smtplib.SMTP_SSL(MAIL["SERVER"]) as smtp:

    print("Logueando en mail... ", end="")
    smtp.login(MAIL["ADDRESS"], MAIL["PASSWORD"])
    print("Listo!")

    # para cada estudiante enviamos su nota
    for _, alum in df_notas.iterrows():

      print(f"- {alum['Nombre Completo']}: {alum['Dirección de correo']}")

      # creamos el mensaje a enviar
      msg = EmailMessage()
      msg['Subject'] = f"Examen"
      msg['From'] = MAIL["ADDRESS"]
      msg['To'] = alum["Dirección de correo"]
      msg['CC'] = MAIL_PROFESOR + ", " + MAIL["ADDRESS"]
      msg.set_content(
        f"Estimada/o {alum['Nombre Completo']},\n"
        f"\n"
        f"Le compartimos la devolución del examen de la materia y su resolución escaneada (adjunta)\n"
        f"\n"
        f"Ejercicio 1\n"
        f"\t○ Nota: {alum['Nota ej1 (30 puntos)']} / 30\n"
        f"\t○ Observaciones: {alum['Observaciones ej1']}\n"    
        f"\n"
        f"Ejercicio 2\n"
        f"\t○ Nota: {alum['Nota ej2 (40 puntos)']} / 40\n"
        f"\t○ Observaciones: {alum['Observaciones ej2']}\n"
        f"\n"
        f"Ejercicio 3\n"
        f"\t○ Nota: {alum['Nota ej3 (30 puntos)']} / 30\n"
        f"\t○ Observaciones: {alum['Observaciones ej3']}\n"
        f"\n"
        f"● Nota del Examen: {alum['Nota Final']} ({alum['Aprobado']})\n"
        f"\n"
        f"Cualquier duda, quedamos a disposición!\n"
        f"\n"
        f"Saludos,\n"
        f"Federico Duca"
      )

      # Los archivos con los examenes presenciales y remotos están en carpetas distintas:
      if alum["Entregó"] == "Remoto":
        path_examen = f"../entregas/remotos/{alum['Nombre'].lower()}-{alum['Apellido'].lower()}.pdf"
      elif alum["Entregó"] == "Presencial":
        path_examen = f"../entregas/presenciales/{alum['Apellido'].lower()}, {alum['Nombre'].lower()}.pdf"

      # abrimos y guardamos el contenido localmente
      with open(path_examen, 'rb') as f:
        msg.add_attachment(f.read(), maintype="application",
                   subtype="octet-stream", filename=f.name.split("/")[-1])

      smtp.send_message(msg)
      
  print("Gracias vuelva prontos")
