#LIBRERIAS NECESARIAS
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import smtplib, getpass, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64

#CUERPO DE LA APP
root=Tk()
root.title("Envio Correo SRT")
root.geometry("320x330")
root.resizable(False,False)
root.iconbitmap(r'C:\Users\lizandro.vivanco\Documents\Visual Code\Aplicaciones\Email.ico')

miFrame=Frame(root)
miFrame.pack()

#VARIABLES
msjto=StringVar()
msjcc=StringVar()
msjsubject=StringVar()
nameshow=StringVar()
filename=StringVar()

#CUENTA CORREO
user = "svc-anovope.notifica@ingrammicro.com"
pasw = "q7XH!V6L"
de = "svc-anovope.notifica@ingrammicro.com"


#FUNCIONES
def saliraplicacion():
    valor=messagebox.askquestion("Salir", "Desea salir de la aplicacion?")

    if valor=="yes":
        root.destroy()

def borrarcampos():
    msjto.set("jhoel.cabello@ingrammicro.com")
    msjcc.set("lizandro.vivanco@ingrammicro.com")
    msjsubject.set("")
    nameshow.set("Seleccionar Archivo...")
    mesagetext.delete(1.0, END)
    filename.set("")

def default():
    msjto.set("jhoel.cabello@ingrammicro.com")
    msjcc.set("lizandro.vivanco@ingrammicro.com")
    msjsubject.set("")
    nameshow.set("Seleccionar Archivo...")
    filename.set("")

def enviarmsj():
    dest =  msjto.get()
    copia = msjcc.get()
    asunto =  msjsubject.get()
    mensaje =   mesagetext.get("1.0", END)
    archivo = filename.get()

    gmail = smtplib.SMTP("smtp.office365.com",587)
    gmail.starttls()
    try:
        gmail.login(user,pasw)
        gmail.set_debuglevel(1)

        rcpt = copia.split(",") + [dest]

        header = MIMEMultipart()
        header['Subject'] = asunto
        header['From'] = de
        header['To'] = dest
        header['Cc'] = copia
    
        mensaje = MIMEText(mensaje, 'html')
        header.attach(mensaje)

        if (os.path.isfile(archivo)):
            adjunto = MIMEBase('application', 'octet-stream')
            adjunto.set_payload(open(archivo, "rb").read())
            encode_base64(adjunto)
            adjunto.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(archivo))
            header.attach(adjunto)

        gmail.sendmail(de,rcpt,header.as_string())

        gmail.quit()

        messagebox.showinfo("Email", "Correo enviado con exito")

    except:
        messagebox.showwarning("Atencion", "Error de envio")
        gmail.quit()


def adjuntararchivo():
    global namefile
    namefile =  filedialog.askopenfilename(title = "Seleccionar Archivo...",filetypes = (("Archivos Excel","*.xlsx"),("Todos los Archivos","*.*")))
    try:
        if namefile == "":
            nameshow.set("Seleccionar Archivo...")
        else:
            nameshow.set(os.path.basename(namefile))
            filename.set(namefile)
            msjsubject.set(os.path.basename(namefile))
    except:
        nameshow.set("Seleccionar Archivo...")

default()
#------------LABELS-----------#
tolabel=Label(miFrame, text="Para:")
tolabel.grid(row=0, column=0, padx=5, pady=5, sticky="e")

cclabel=Label(miFrame, text="Cc:")
cclabel.grid(row=1, column=0, padx=5, pady=5, sticky="e")

subjectlabel=Label(miFrame, text="Asunto:")
subjectlabel.grid(row=2, column=0, padx=5, pady=5, sticky="e")

adjuntarlabel=Label(miFrame, text="Adjuntar:")
adjuntarlabel.grid(row=3, column=0, padx=5, pady=5, sticky="e")

mesagelabel=Label(miFrame, text="Mensaje:")
mesagelabel.grid(row=4, column=0, padx=5, pady=5, sticky="n")

#---------ENTRY-------
toentry=Entry(miFrame, width=40, textvariable=msjto,state='disabled')
toentry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)

ccentry=Entry(miFrame, width=40, textvariable=msjcc,state='disabled')
ccentry.grid(row=1, column=1, padx=5, pady=5, columnspan=2)

subjectentry=Entry(miFrame, width=40, textvariable=msjsubject)
subjectentry.grid(row=2, column=1, padx=5, pady=5, columnspan=2)

adjuntarbutton=Button(miFrame, width=30, textvariable=nameshow, command=adjuntararchivo)
adjuntarbutton.grid(row=3,column=1, sticky="w", padx=5, pady=5)

mesagetext=Text(miFrame, width=32, height=10, font=("Arial",9), wrap=WORD)
mesagetext.grid(row=4, column=1, padx=5, pady=5)
scrollvert=Scrollbar(miFrame, command=mesagetext.yview)
scrollvert.grid(row=4, column=2, sticky="nsew")
mesagetext.config(yscrollcommand=scrollvert.set)

miFrame2=Frame(root, width=600, height=1000)
miFrame2.pack()

sendbutton=Button(miFrame2, width=8, text="Enviar", command=enviarmsj)
sendbutton.grid(row=0,column=0, sticky="e", padx=5, pady=5)

deletebutton=Button(miFrame2, width=8, text="Borrar", command=borrarcampos)
deletebutton.grid(row=0,column=1, sticky="e", padx=5, pady=5)

defaultbutton=Button(miFrame2, width=8, text="Defecto", command=default)
defaultbutton.grid(row=0,column=2, sticky="e", padx=5, pady=5)

exitbutton=Button(miFrame2, width=8, text="Salir", command=saliraplicacion)
exitbutton.grid(row=0,column=3, sticky="e", padx=5, pady=5)

root.mainloop()