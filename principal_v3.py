# principal_v3.py
# Versión mejorada con:
# - Envío de correos con imagen embebida
# - Login con alineación simétrica
# - Ordenamiento ascendente/descendente en columnas
# - Campo de búsqueda con botones Buscar/Limpiar

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from ldap3 import Server, Connection, ALL
import datetime

# =====================================
# CONFIGURACIONES GLOBALES
# =====================================
SMTP_SERVER = "smtp.capual.cl"
SMTP_PORT = 587
SMTP_REMITENTE = "agente.ti2@capual.cl"
SMTP_PASSWORD = "L18835654-0"
APP_CREDITOS = "Capual Password Alert"

# =====================================
# FUNCIONES BASE
# =====================================
def centrar_ventana(win, width, height):
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")

def confirmar_y_cerrar(win):
    if messagebox.askokcancel("Salir", "¿Deseas cerrar la aplicación?"):
        win.destroy()

def conectar_ldap(usuario, contrasena):
    server = Server('SRV_DC01_NEW.capual.cl', get_info=ALL)
    conn = Connection(server, user=f'{usuario}@capual.cl', password=contrasena, authentication='SIMPLE', auto_bind=True)
    return conn

# =====================================
# ENVÍO DE CORREOS (con imagen embebida)
# =====================================
def enviar_correos_lista(usuarios):
    if not usuarios:
        messagebox.showwarning("Aviso", "No hay usuarios seleccionados para enviar correos.")
        return False

    confirm = messagebox.askyesno("Confirmar envío", f"¿Desea enviar correos a {len(usuarios)} usuarios seleccionados?")
    if not confirm:
        return False

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_REMITENTE, SMTP_PASSWORD)
        enviado_count = 0
        ruta_imagen = "img_teclas.png"

        for u in usuarios:
            if u.get("correo") and "@" in u.get("correo"):
                msg = MIMEMultipart("related")
                msg["From"] = SMTP_REMITENTE
                msg["To"] = u["correo"]
                msg["Subject"] = "⚠️ Aviso: Tu contraseña está próxima a expirar"

                html_body = f"""
                <html>
                <body style="font-family:Segoe UI, sans-serif; color:#333;">
                    <p>Estimado/a <b>{u['nombre']}</b>,</p>
                    <p>Tu contraseña expira en <b>{u['dias']} días</b> (el {u['expira']}).<br>
                    Por favor, actualízala antes de que caduque para evitar bloqueos de acceso.</p>
                    <p><b>Para cambiar tu contraseña:</b><br>
                    Presiona <i>Ctrl + Alt + Supr</i> y selecciona la opción "Cambiar contraseña".</p>
                    <p style="text-align:center;">
                        <img src="cid:img_teclas" alt="Instrucciones Ctrl+Alt+Supr" width="420">
                    </p>
                    <p>Si tienes problemas, comunícate con:<br>
                    - Eduardo L. (Nexo 4006)<br>
                    - Ignacio C. (Nexo 4018)<br>
                    Departamento de Servicios TI</p>
                    <p>🏢 Departamento: {u['departamento'] or 'No especificado'}<br>
                    👤 Usuario: {u['usuario']}</p>
                    <p>Saludos cordiales,<br><b>Departamento de Soporte TI</b><br>
                    Capual - Cooperativa de Ahorro y Crédito</p>
                </body></nhtml>"""

                msg.attach(MIMEText(html_body, "html"))

                try:
                    with open(ruta_imagen, "rb") as f:
                        img = MIMEImage(f.read())
                        img.add_header("Content-ID", "<img_teclas>")
                        img.add_header("Content-Disposition", "inline", filename="img_teclas.png")
                        msg.attach(img)
                except FileNotFoundError:
                    print(f"⚠️ Imagen {ruta_imagen} no encontrada.")

                server.send_message(msg)
                enviado_count += 1

        server.quit()
        messagebox.showinfo("Envío finalizado", f"Correos enviados: {enviado_count}")
        return True

    except Exception as e:
        messagebox.showerror("Error envío", f"No se pudieron enviar los correos:\n{e}")
        return False

# =====================================
# INTERFAZ: LOGIN
# =====================================
def ventana_login():
    login_win = tk.Tk()
    login_win.title(f"Iniciar sesión - {APP_CREDITOS}")
    centrar_ventana(login_win, 420, 230)
    login_win.resizable(False, False)
    login_win.protocol("WM_DELETE_WINDOW", lambda: confirmar_y_cerrar(login_win))

    frm = ttk.Frame(login_win, padding=12)
    frm.pack(fill="both", expand=True)

    frm.columnconfigure(0, weight=0)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="Usuario: ").grid(row=0, column=0, sticky="e", pady=(8,4), padx=(4,4))
    usuario_entry = ttk.Entry(frm, width=35)
    usuario_entry.grid(row=0, column=1, pady=(8,4), padx=(4,4))

    ttk.Label(frm, text="Contraseña: ").grid(row=1, column=0, sticky="e", pady=(4,8), padx=(4,4))
    pass_entry = ttk.Entry(frm, width=35, show="*")
    pass_entry.grid(row=1, column=1, pady=(4,8), padx=(4,4))

    status_lbl = ttk.Label(frm, text="")
    status_lbl.grid(row=2, column=0, columnspan=2, pady=(8,4))

    btn_frame = ttk.Frame(frm)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=(10,4))

    def intentar_login():
        user = usuario_entry.get().strip()
        pw = pass_entry.get().strip()
        if not user or not pw:
            messagebox.showwarning("Datos incompletos", "Debes ingresar usuario y contraseña.")
            return
        try:
            status_lbl.config(text="Conectando al AD...")
            login_win.update_idletasks()
            conn = conectar_ldap(user, pw)
            login_win.destroy()
            ventana_principal(conn)
        except Exception:
            messagebox.showerror("Error", "Credenciales inválidas o no se pudo conectar al AD.")
            usuario_entry.focus()
            status_lbl.config(text="")

    ttk.Button(btn_frame, text="Iniciar sesión", command=intentar_login).grid(row=0, column=0, padx=6)
    ttk.Button(btn_frame, text="Salir", command=lambda: confirmar_y_cerrar(login_win)).grid(row=0, column=1, padx=6)

    login_win.mainloop()

# =====================================
# FUNCIONES DE CONSULTA (simuladas para ejemplo)
# =====================================
def consultar_usuarios(conn):
    hoy = datetime.date.today()
    return [
        {"usuario": "jrojas", "nombre": "Juan Rojas", "correo": "jrojas@capual.cl", "departamento": "Contabilidad", "dias": 3, "expira": (hoy + datetime.timedelta(days=3)).strftime('%d/%m/%Y')},
        {"usuario": "mvera", "nombre": "María Vera", "correo": "mvera@capual.cl", "departamento": "TI", "dias": 7, "expira": (hoy + datetime.timedelta(days=7)).strftime('%d/%m/%Y')},
        {"usuario": "fgarcia", "nombre": "Felipe García", "correo": "fgarcia@capual.cl", "departamento": "Créditos", "dias": 2, "expira": (hoy + datetime.timedelta(days=2)).strftime('%d/%m/%Y')}
    ]

# =====================================
# INTERFAZ PRINCIPAL
# =====================================
def ventana_principal(conn):
    main_win = tk.Tk()
    main_win.title(APP_CREDITOS)
    centrar_ventana(main_win, 900, 600)
    main_win.protocol("WM_DELETE_WINDOW", lambda: confirmar_y_cerrar(main_win))

    frame = ttk.Frame(main_win, padding=10)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text="Usuarios próximos a expirar:", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))

    # --- buscador ---
    buscador_frame = ttk.Frame(frame)
    buscador_frame.pack(fill="x", pady=(0, 8))
    ttk.Label(buscador_frame, text="🔍 Buscar:").pack(side="left", padx=(0, 6))
    entrada_busqueda = ttk.Entry(buscador_frame, width=40)
    entrada_busqueda.pack(side="left", padx=(0, 6))

    cols = ("Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)

    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=150, anchor="w")
    tree.pack(fill="both", expand=True)

    usuarios = consultar_usuarios(conn)
    for u in usuarios:
        tree.insert("", "end", values=(u["usuario"], u["nombre"], u["correo"], u["departamento"], u["dias"], u["expira"]))

    # --- ordenamiento ---
    def ordenar_tabla(col, reverse=False):
        datos = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            datos.sort(key=lambda t: (int(t[0]) if t[0].isdigit() else t[0].lower()), reverse=reverse)
        except Exception:
            datos.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for idx, (val, k) in enumerate(datos):
            tree.move(k, '', idx)
        tree.heading(col, command=lambda: ordenar_tabla(col, not reverse))

    for col in cols:
        tree.heading(col, text=col, command=lambda c=col: ordenar_tabla(c))

    # --- búsqueda ---
    def filtrar_tabla():
        texto = entrada_busqueda.get().strip().lower()
        for item in tree.get_children():
            valores = tree.item(item, "values")
            visible = any(texto in str(v).lower() for v in valores)
            tree.detach(item) if not visible else tree.reattach(item, '', 'end')

    def limpiar_filtro():
        entrada_busqueda.delete(0, tk.END)
        for item in tree.get_children():
            tree.reattach(item, '', 'end')

    ttk.Button(buscador_frame, text="Buscar", command=filtrar_tabla).pack(side="left", padx=(0, 6))
    ttk.Button(buscador_frame, text="Limpiar", command=limpiar_filtro).pack(side="left")

    ttk.Button(frame, text="Enviar avisos", command=lambda: enviar_correos_lista(usuarios)).pack(pady=10)

    main_win.mainloop()

if __name__ == "__main__":
    ventana_login()