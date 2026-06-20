import tkinter as tk
from tkinter import simpledialog
from tkinter import ttk
import json
import os
import sys
import win32com.client as win32
import datetime

CONFIG_FILE = "config.json"

# Determinar la ruta base: si el programa está congelado (ejecutable) o en modo script.
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE_PATH = os.path.join(base_path, CONFIG_FILE)

def load_config():
    if os.path.exists(CONFIG_FILE_PATH):
        with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {}
    config.setdefault("destinatarios", [])
    config.setdefault("asunto", "")
    config.setdefault("cuerpo", "")
    config.setdefault("auto_close", True)
    config.setdefault("auto_close_delay", 60)
    return config

def update_status(message, color="black"):
    status_label.config(text=message, fg=color)

def save_config():
    config_to_save = {
        "destinatarios": list(listbox_destinatarios.get(0, tk.END)),
        "asunto": entry_asunto.get(),
        "cuerpo": text_cuerpo.get("1.0", tk.END).strip(),
        "auto_close": auto_close_var.get(),
        "auto_close_delay": int(auto_close_delay_var.get())
    }
    try:
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(config_to_save, f, indent=4, ensure_ascii=False)
        update_status("Configuración guardada.", "green")
    except Exception as e:
        update_status(f"Error al guardar la configuración: {e}", "red")

def agregar_destinatario():
    nuevo = simpledialog.askstring("Agregar destinatario", "Ingrese el correo del destinatario:")
    if nuevo:
        listbox_destinatarios.insert(tk.END, nuevo.strip())
        update_status("Destinatario agregado.", "green")

def eliminar_destinatario():
    seleccion = listbox_destinatarios.curselection()
    if not seleccion:
        update_status("Seleccione un destinatario para eliminar.", "red")
        return
    for index in seleccion[::-1]:
        listbox_destinatarios.delete(index)
    update_status("Destinatario(s) eliminado(s).", "green")

def obtener_cuentas():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        accounts = outlook.Session.Accounts
        cuentas = []
        for account in accounts:
            cuentas.append(account.SmtpAddress)
        return cuentas
    except Exception as e:
        update_status("Error al obtener cuentas: " + str(e))
        return []

def enviar_correo():
    destinatarios = [d.strip() for d in listbox_destinatarios.get(0, tk.END) if d.strip()]
    if not destinatarios:
        update_status("Agregue al menos un destinatario.", "red")
        return

    asunto = entry_asunto.get().strip()
    cuerpo = text_cuerpo.get("1.0", tk.END).strip()

    now = datetime.datetime.now()
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    month = now.month
    year = now.year

    if month == 1:
        mes_anterior = 12
        year_anterior = year - 1
    else:
        mes_anterior = month - 1
        year_anterior = year

    month_name_anterior = meses[mes_anterior].capitalize()
    year_num = str(year_anterior)

    # Reemplazar placeholders en asunto y cuerpo
    asunto = asunto.replace("[Mes anterior en letras]", month_name_anterior).replace("[año en numero]", year_num)
    cuerpo = cuerpo.replace("[Mes anterior en letras]", month_name_anterior).replace("[año en numero]", year_num)

    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = asunto
        mail.Body = cuerpo

        selected_account = combobox_cuenta.get()
        if selected_account:
            account_found = False
            for account in outlook.Session.Accounts:
                if account.SmtpAddress.lower() == selected_account.lower():
                    mail.SendUsingAccount = account
                    account_found = True
                    break
            if not account_found:
                update_status("No se encontró la cuenta de envío seleccionada.", "red")
                return

        filtered_destinatarios = [d for d in destinatarios if d.lower() != selected_account.lower()]
        mail.To = "; ".join(filtered_destinatarios)

        mail.Send()
        update_status("Correo enviado exitosamente.", "green")

        if auto_close_var.get():
            delay_segundos = int(auto_close_delay_var.get())
            iniciar_countdown(delay_segundos)

    except Exception as e:
        update_status(f"No se pudo enviar el correo. Error: {e}", "red")

def iniciar_countdown(segundos):
    def countdown():
        nonlocal segundos
        if segundos > 0:
            update_status(f"Correo enviado. Cerrando en {segundos} segundos...", "green")
            segundos -= 1
            root.after(1000, countdown)
        else:
            btn_salir.invoke()
    countdown()

def salir():
    root.destroy()

# Cargar configuración inicial
config = load_config()

# Interfaz gráfica
root = tk.Tk()
root.title("Enviar Correo Outlook 365")

frame_destinatarios = tk.Frame(root)
frame_destinatarios.pack(pady=(10, 0), fill=tk.X, padx=10)

label_destinatarios = tk.Label(frame_destinatarios, text="Destinatarios:")
label_destinatarios.pack(anchor="w")

listbox_destinatarios = tk.Listbox(frame_destinatarios, width=50, height=5)
for destinatario in config.get("destinatarios", []):
    listbox_destinatarios.insert(tk.END, destinatario)
listbox_destinatarios.pack(pady=5)

frame_buttons_dest = tk.Frame(frame_destinatarios)
frame_buttons_dest.pack()

btn_agregar = tk.Button(frame_buttons_dest, text="Agregar", width=15, command=agregar_destinatario)
btn_agregar.pack(side=tk.LEFT, padx=5)

btn_eliminar = tk.Button(frame_buttons_dest, text="Eliminar", width=15, command=eliminar_destinatario)
btn_eliminar.pack(side=tk.LEFT, padx=5)

label_asunto = tk.Label(root, text="Asunto:")
label_asunto.pack(pady=(10, 0))
entry_asunto = tk.Entry(root, width=50)
entry_asunto.insert(0, config.get("asunto", ""))
entry_asunto.pack(pady=5)

label_cuerpo = tk.Label(root, text="Cuerpo del correo:")
label_cuerpo.pack(pady=(10, 0))
text_cuerpo = tk.Text(root, width=50, height=10)
text_cuerpo.insert("1.0", config.get("cuerpo", ""))
text_cuerpo.pack(pady=5)

frame_cuenta = tk.LabelFrame(root, text="Cuenta de Envío")
frame_cuenta.pack(padx=10, pady=5, fill="both")

label_cuenta = tk.Label(frame_cuenta, text="Seleccione la cuenta de envío:")
label_cuenta.pack(anchor="w", padx=10, pady=5)

cuentas = obtener_cuentas()
combobox_cuenta = ttk.Combobox(frame_cuenta, values=cuentas, state="readonly")
if cuentas:
    combobox_cuenta.current(0)
combobox_cuenta.pack(anchor="w", padx=10, pady=5)

frame_auto_close = tk.LabelFrame(root, text="Configuración de Cierre Automático")
frame_auto_close.pack(padx=10, pady=5, fill="both")

auto_close_var = tk.BooleanVar(value=config.get("auto_close", True))
chk_auto_close = tk.Checkbutton(frame_auto_close, text="Cerrar automáticamente la aplicación tras enviar el correo", variable=auto_close_var)
chk_auto_close.pack(anchor="w", padx=10, pady=5)

auto_close_delay_var = tk.StringVar(value=str(config.get("auto_close_delay", 60)))
label_delay = tk.Label(frame_auto_close, text="Tiempo de espera para cierre automático (segundos):")
label_delay.pack(anchor="w", padx=10)
entry_delay = tk.Entry(frame_auto_close, textvariable=auto_close_delay_var, width=10)
entry_delay.pack(anchor="w", padx=10, pady=5)

frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)

btn_enviar = tk.Button(frame_buttons, text="Enviar", width=20, command=enviar_correo)
btn_enviar.pack(side=tk.LEFT, padx=5)

btn_guardar = tk.Button(frame_buttons, text="Guardar configuración", width=20, command=save_config)
btn_guardar.pack(side=tk.LEFT, padx=5)

btn_salir = tk.Button(frame_buttons, text="Salir", width=20, command=salir)
btn_salir.pack(side=tk.LEFT, padx=5)

status_label = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

root.after(1000, btn_enviar.invoke)

root.mainloop()
