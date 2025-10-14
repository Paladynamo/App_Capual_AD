from ldap3 import Server, Connection, ALL, SUBTREE
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog

# ==============================
# CONFIGURACIÓN DE CONEXIÓN LDAP
# ==============================
AD_SERVER = 'ldaps://SRV_DC01_NEW.capual.cl'
BASE_DN = 'DC=capual,DC=cl'

# ==============================
# FUNCIONES DE INTERFAZ
# ==============================

def centrar_ventana(ventana, ancho, alto):
    """Centrar una ventana en la pantalla"""
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


def seleccionar_usuario():
    """Pantalla inicial: selección de usuario"""
    def continuar():
        seleccion = combo_usuarios.get()
        if not seleccion:
            messagebox.showwarning("Atención", "Por favor selecciona un usuario para continuar.")
            return
        global AD_USER
        AD_USER = f"{seleccion}@capual.cl"
        root_sel.destroy()

    root_sel = tk.Tk()
    root_sel.title("Bienvenido - App creada por Eduardo 'PaladynamoX' Lizama C.")
    centrar_ventana(root_sel, 400, 240)
    root_sel.resizable(False, False)

    label_bienvenida = ttk.Label(root_sel, text="👋 ¡Bienvenido al sistema de supervisión!", font=("Segoe UI", 11))
    label_bienvenida.pack(pady=10)

    label_creditos = ttk.Label(root_sel, text="App creada por Eduardo 'PaladynamoX' Lizama C.\nVersión 1.0.0 - Año 2025", font=("Segoe UI", 9))
    label_creditos.pack(pady=5)

    label_instruccion = ttk.Label(root_sel, text="Selecciona el usuario que utilizará el programa:")
    label_instruccion.pack(pady=5)

    usuarios = ["agente.ti1", "agente.ti2", "agente.infra", "agente.redes", "agente.plataformas"]
    combo_usuarios = ttk.Combobox(root_sel, values=usuarios, state="readonly", width=25)
    combo_usuarios.pack(pady=5)
    combo_usuarios.set("agente.ti2")

    btn_continuar = ttk.Button(root_sel, text="Continuar", command=continuar)
    btn_continuar.pack(pady=10)

    root_sel.mainloop()


def pedir_dias_aviso():
    """Pide cuántos días faltantes se quieren consultar"""
    temp_root = tk.Tk()
    temp_root.withdraw()
    dias = simpledialog.askinteger(
        "Filtro de días",
        "¿Cuántos días antes del vencimiento deseas mostrar?",
        minvalue=1,
        maxvalue=90,
        parent=temp_root
    )
    temp_root.destroy()
    return dias if dias else 10


def conectar_ad():
    """Intenta conectar con AD y vuelve al inicio si hay error"""
    while True:
        seleccionar_usuario()

        temp_root = tk.Tk()
        temp_root.withdraw()
        AD_PASSWORD = simpledialog.askstring(
            "Credenciales AD",
            f"🔑 Ingresa la contraseña de {AD_USER}:",
            show='*',
            parent=temp_root
        )
        temp_root.destroy()

        try:
            server = Server(AD_SERVER, get_info=ALL)
            conn = Connection(server, user=AD_USER, password=AD_PASSWORD, authentication='SIMPLE', auto_bind=True)
            print("✅ Conexión establecida con Active Directory")
            return conn
        except Exception:
            messagebox.showerror("Error de autenticación", "❌ Contraseña incorrecta o credenciales inválidas.\nIntenta nuevamente.")


# ==============================
# CONSULTA DE USUARIOS
# ==============================
def consultar_usuarios(conn, dias_aviso):
    conn.search(
        BASE_DN,
        "(&"
        "(objectCategory=person)"
        "(objectClass=user)"
        "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        "(!(userAccountControl:1.2.840.113556.1.4.803:=65536))"
        "(!(sAMAccountName=*$))"
        "(!(sAMAccountName=Administrador))"
        "(!(sAMAccountName=Agente.*))"
        ")",
        SUBTREE,
        attributes=['sAMAccountName', 'displayName', 'mail', 'msDS-UserPasswordExpiryTimeComputed', 'department']
    )

    now = datetime.now()
    usuarios_por_vencer = []

    for entry in conn.entries:
        expiry_raw = entry['msDS-UserPasswordExpiryTimeComputed'].value
        if not expiry_raw or int(expiry_raw) <= 0:
            continue
        try:
            expiry_date = datetime.fromtimestamp(int(expiry_raw) / 1e7 - 11644473600)
            dias_restantes = (expiry_date - now).days
            if 0 <= dias_restantes <= dias_aviso:
                usuarios_por_vencer.append({
                    "usuario": str(entry['sAMAccountName']),
                    "nombre": str(entry['displayName']),
                    "correo": str(entry['mail']),
                    "departamento": str(entry['department']),
                    "dias": dias_restantes,
                    "expira": expiry_date.strftime("%d/%m/%Y %H:%M")
                })
        except Exception as e:
            print(f"⚠️ No se pudo procesar {entry['sAMAccountName']}: {e}")

    return usuarios_por_vencer


# ==============================
# FUNCIÓN PARA ENVIAR CORREOS
# ==============================
def enviar_correos(usuarios_por_vencer, refrescar_callback):
    if not usuarios_por_vencer:
        messagebox.showinfo("Información", "No hay usuarios próximos a vencer su contraseña.")
        return

    respuesta = messagebox.askyesno("Confirmar envío", "¿Desea enviar los mensajes de aviso por correo?")
    if respuesta:
        remitente = "printservice@capual.cl"
        remitente_pass = "PSD34$/srvc123."
        smtp_server = "smtp.office365.com"
        smtp_port = 587

        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(remitente, remitente_pass)

            for usuario in usuarios_por_vencer:
                if usuario["correo"] and "@" in usuario["correo"]:
                    mensaje = MIMEMultipart()
                    mensaje["From"] = remitente
                    mensaje["To"] = usuario["correo"]
                    mensaje["Subject"] = "⚠️ Aviso: Tu contraseña está próxima a expirar"

                    cuerpo = f"""
Estimado/a {usuario["nombre"]},

Tu contraseña expira en {usuario["dias"]} días (el {usuario["expira"]}).
Por favor, actualízala antes de que caduque para evitar bloqueos de acceso.

🏢 Departamento: {usuario["departamento"] or "No especificado"}
👤 Usuario: {usuario["usuario"]}

Saludos cordiales,
Departamento de Soporte TI
Capual - Cooperativa de Ahorro y Crédito
                    """
                    mensaje.attach(MIMEText(cuerpo, "plain"))
                    server.send_message(mensaje)

            server.quit()
            messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo enviar los correos:\n{e}")

    nueva_busqueda = messagebox.askyesno("Nueva búsqueda", "¿Deseas aplicar un nuevo filtro de días?")
    if nueva_busqueda:
        dias_aviso = pedir_dias_aviso()
        refrescar_callback(dias_aviso)
    else:
        messagebox.showinfo("Despedida", "Gracias por usar esta aplicación.\nApp creada por Eduardo 'PaladynamoX' Lizama C.\nVersión 1.0.0 - Año 2025")
        root.destroy()


# ==============================
# INTERFAZ PRINCIPAL
# ==============================
def mostrar_ventana_principal(conn, dias_aviso):
    global root, tabla

    usuarios_por_vencer = consultar_usuarios(conn, dias_aviso)

    def refrescar_tabla(nuevos_dias):
        for row in tabla.get_children():
            tabla.delete(row)
        nuevos_usuarios = consultar_usuarios(conn, nuevos_dias)
        for u in nuevos_usuarios:
            tabla.insert("", "end", values=(
                u["usuario"], u["nombre"], u["correo"], u["departamento"], u["dias"], u["expira"]
            ))
        tabla.usuarios = nuevos_usuarios

    def sort_by(col, descending):
        data = [(tabla.set(child, col), child) for child in tabla.get_children('')]
        if col in ["Días restantes"]:
            data.sort(key=lambda t: int(t[0]), reverse=descending)
        else:
            data.sort(reverse=descending)
        for idx, item in enumerate(data):
            tabla.move(item[1], '', idx)
        tabla.heading(col, command=lambda: sort_by(col, not descending))

    root = tk.Tk()
    root.title("Usuarios con contraseña próxima a expirar - App PaladynamoX v1.0.0")
    centrar_ventana(root, 950, 420)

    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    cols = ("Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
    tabla = ttk.Treeview(frame, columns=cols, show='headings')

    for col in cols:
        tabla.heading(col, text=col, command=lambda c=col: sort_by(c, False))
        tabla.column(col, width=150)

    for u in usuarios_por_vencer:
        tabla.insert("", "end", values=(
            u["usuario"], u["nombre"], u["correo"], u["departamento"], u["dias"], u["expira"]
        ))

    tabla.usuarios = usuarios_por_vencer
    tabla.pack(fill="both", expand=True)

    btn_enviar = ttk.Button(root, text="📧 Enviar avisos por correo",
                            command=lambda: enviar_correos(tabla.usuarios, refrescar_tabla))
    btn_enviar.pack(pady=10)

    if not usuarios_por_vencer:
        messagebox.showinfo("Sin resultados", "No hay usuarios próximos a expirar su contraseña.")

    root.mainloop()


# ==============================
# EJECUCIÓN PRINCIPAL
# ==============================
if __name__ == "__main__":
    conn = conectar_ad()
    dias_aviso = pedir_dias_aviso()
    mostrar_ventana_principal(conn, dias_aviso)