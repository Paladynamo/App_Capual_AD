# principal_v2.py
from ldap3 import Server, Connection, ALL, SUBTREE
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import sys
import matplotlib
matplotlib.use("Agg")  # backend para crear figuras sin GUI directa
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ==============================
# CONFIGURACIÓN (NO CAMBIAR si no es necesario)
# ==============================
AD_SERVER = 'ldaps://SRV_DC01_NEW.capual.cl'
BASE_DN = 'DC=capual,DC=cl'

SMTP_REMITENTE = "printservice@capual.cl"
SMTP_PASSWORD = "PSD34$/srvc123."
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

APP_CREDITOS = "App creada por Eduardo 'PaladynamoX' Lizama C.\nVersión 2.0.0 - Año 2025"

# ==============================
# UTILIDADES UI
# ==============================
def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

def despedida_final():
    messagebox.showinfo("Despedida", f"Gracias por usar esta aplicación.\n{APP_CREDITOS}")
    try:
        root_all.destroy()
    except Exception:
        pass
    sys.exit(0)

def confirmar_y_cerrar(ventana):
    if messagebox.askyesno("Confirmar salida", "¿Deseas salir del programa?"):
        despedida_final()
    else:
        # si se cancela, no hacer nada (la ventana permanece)
        return

# ==============================
# FUNCIONES LDAP / CONSULTAS
# ==============================
def conectar_ldap(username, password):
    # username puede venir como 'usuario' o 'usuario@dominio'
    user = username
    if "@" not in user:
        user = f"{user}@capual.cl"
    try:
        server = Server(AD_SERVER, get_info=ALL)
        conn = Connection(server, user=user, password=password, authentication='SIMPLE', auto_bind=True)
        return conn
    except Exception as e:
        raise

def msds_to_datetime(msds_value):
    # msDS-UserPasswordExpiryTimeComputed viene en ticks (100-ns desde 1601)
    try:
        if not msds_value or int(msds_value) <= 0:
            return None
        expiry_date = datetime.fromtimestamp(int(msds_value) / 1e7 - 11644473600)
        return expiry_date
    except Exception:
        return None

def consultar_usuarios(conn):
    # Consulta todos los usuarios relevantes (no muestra cuentas de equipo, admin ni los que empiezan con Agente.)
    filter_query = (
        "(&"
        "(objectCategory=person)"
        "(objectClass=user)"
        "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"
        "(!(userAccountControl:1.2.840.113556.1.4.803:=65536))"
        "(!(sAMAccountName=*$))"
        "(!(sAMAccountName=Administrador))"
        "(!(sAMAccountName=Agente.*))"
        ")"
    )
    attributes = ['sAMAccountName', 'displayName', 'mail', 'msDS-UserPasswordExpiryTimeComputed', 'department']
    conn.search(BASE_DN, filter_query, SUBTREE, attributes=attributes)
    results = []
    now = datetime.now()
    for entry in conn.entries:
        try:
            sAM = str(entry['sAMAccountName']) if entry['sAMAccountName'].value else ""
            display = str(entry['displayName']) if entry['displayName'].value else sAM
            mail = str(entry['mail']) if entry['mail'].value else ""
            dept = str(entry['department']) if entry['department'].value else ""
            expiry_raw = entry['msDS-UserPasswordExpiryTimeComputed'].value
            expiry_dt = msds_to_datetime(expiry_raw)
            if not expiry_dt:
                continue
            dias_restantes = (expiry_dt - now).days
            results.append({
                "usuario": sAM,
                "nombre": display,
                "correo": mail,
                "departamento": dept,
                "dias": dias_restantes,
                "expira": expiry_dt.strftime("%d/%m/%Y %H:%M")
            })
        except Exception as e:
            # ignorar entradas mal formadas
            continue
    return results

# ==============================
# ENVÍO DE CORREOS
# ==============================
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
        for u in usuarios:
            if u.get("correo") and "@" in u.get("correo"):
                msg = MIMEMultipart()
                msg["From"] = SMTP_REMITENTE
                msg["To"] = u["correo"]
                msg["Subject"] = "⚠️ Aviso: Tu contraseña está próxima a expirar"

                body = f"""Estimado/a {u['nombre']},

Tu contraseña expira en {u['dias']} días (el {u['expira']}).
Por favor, actualízala antes de que caduque para evitar bloqueos de acceso.

Para cambiar tu contraseña:
Presiona Ctrl + Alt + Supr y selecciona la opción "Cambiar contraseña".

Si tienes problemas, comunícate con:
- Eduardo L. (Nexo 4006)
- Ignacio C. (Nexo 4018)
Departamento de Servicios TI

🏢 Departamento: {u['departamento'] or 'No especificado'}
👤 Usuario: {u['usuario']}

Saludos cordiales,
Departamento de Soporte TI
Capual - Cooperativa de Ahorro y Crédito
"""
                msg.attach(MIMEText(body, "plain"))
                server.send_message(msg)
                enviado_count += 1
        server.quit()
        messagebox.showinfo("Envío finalizado", f"Correos enviados: {enviado_count}")
        return True
    except Exception as e:
        messagebox.showerror("Error envío", f"No se pudieron enviar los correos:\n{e}")
        return False

# ==============================
# VENTANAS / FLUJO
# ==============================
# Mantener referencia global para cierre limpio
root_all = None

def ventana_login():
    global root_all
    login_win = tk.Tk()
    root_all = login_win
    login_win.title(f"Iniciar sesión - {APP_CREDITOS}")
    centrar_ventana(login_win, 420, 230)
    login_win.resizable(False, False)
    login_win.protocol("WM_DELETE_WINDOW", lambda: confirmar_y_cerrar(login_win))

    frm = ttk.Frame(login_win, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Usuario (sAMAccountName o user@dominio):").pack(anchor="w", pady=(6,0))
    usuario_entry = ttk.Entry(frm, width=40)
    usuario_entry.pack(pady=2)
    usuario_entry.focus()

    ttk.Label(frm, text="Contraseña:").pack(anchor="w", pady=(8,0))
    pass_entry = ttk.Entry(frm, width=40, show="*")
    pass_entry.pack(pady=2)

    status_lbl = ttk.Label(frm, text="")
    status_lbl.pack(pady=6)

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
            # si conecta, cerrar login y abrir menu principal
            login_win.destroy()
            ventana_principal(conn)
        except Exception:
            messagebox.showerror("Error", "Credenciales inválidas o no se pudo conectar al AD. Intenta nuevamente.")
            usuario_entry.focus()
            status_lbl.config(text="")

    btn_frame = ttk.Frame(frm)
    btn_frame.pack(pady=8)
    ttk.Button(btn_frame, text="Iniciar sesión", command=intentar_login).grid(row=0, column=0, padx=6)
    ttk.Button(btn_frame, text="Salir", command=lambda: confirmar_y_cerrar(login_win)).grid(row=0, column=1, padx=6)

    login_win.mainloop()

def ventana_principal(conn):
    global root_all
    main_win = tk.Tk()
    root_all = main_win
    main_win.title(f"Menú Principal - {APP_CREDITOS}")
    centrar_ventana(main_win, 520, 260)
    main_win.protocol("WM_DELETE_WINDOW", lambda: confirmar_y_cerrar(main_win))

    frm = ttk.Frame(main_win, padding=20)
    frm.pack(expand=True, fill="both")

    ttk.Label(frm, text="Seleccione una opción", font=("Segoe UI", 12)).pack(pady=(0,12))

    btn1 = ttk.Button(frm, text="📋 Usuarios próximos a expirar", width=36,
                      command=lambda: abrir_usuarios_proximos(main_win, conn))
    btn1.pack(pady=8)

    btn2 = ttk.Button(frm, text="📊 Dashboard de contraseñas (gráfico circular)", width=36,
                      command=lambda: abrir_dashboard(main_win, conn))
    btn2.pack(pady=8)

    ttk.Label(frm, text=APP_CREDITOS, font=("Segoe UI", 9)).pack(side="bottom", pady=(12,0))

    main_win.mainloop()

# ---- Ventana: Usuarios próximos a expirar ----
def abrir_usuarios_proximos(parent_win, conn):
    # Cerrar parent (o esconder) y abrir ventana nueva
    parent_win.withdraw()
    win = tk.Toplevel()
    win.title("Usuarios próximos a expirar")
    centrar_ventana(win, 980, 520)
    win.protocol("WM_DELETE_WINDOW", lambda: on_close_subwindow(win, parent_win))
    frame = ttk.Frame(win, padding=10)
    frame.pack(fill="both", expand=True)

    # pedir días
    dias = simpledialog.askinteger("Filtro de días", "¿Cuántos días antes del vencimiento deseas mostrar?", minvalue=1, maxvalue=180, parent=win)
    if not dias:
        dias = 10

    # obtener todos los usuarios y filtrar por dias
    all_users = consultar_usuarios(conn)
    filtered = [u for u in all_users if 0 <= u["dias"] <= dias]

    # tabla con columna de selección (toggle)
    cols = ("Sel", "Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
    tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="none", height=18)
    for c in cols:
        tree.heading(c, text=c)
        if c == "Sel":
            tree.column(c, width=60, anchor="center")
        elif c == "Días restantes":
            tree.column(c, width=110, anchor="center")
        else:
            tree.column(c, width=150, anchor="w")
    tree.pack(side="top", fill="both", expand=True)

    # uso de dict para marcar selección
    seleccion = {}  # item_id -> bool
    item_to_user = {}  # item_id -> user dict

    def insertar_datos():
        for row in tree.get_children():
            tree.delete(row)
        seleccion.clear()
        item_to_user.clear()
        for u in filtered:
            vals = ("", u["usuario"], u["nombre"], u["correo"], u["departamento"], str(u["dias"]), u["expira"])
            iid = tree.insert("", "end", values=vals)
            seleccion[iid] = False
            item_to_user[iid] = u

    insertar_datos()

    # Toggle al hacer doble click o Enter sobre fila
    def toggle_selection(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        seleccion[item] = not seleccion[item]
        tree.set(item, "Sel", "✓" if seleccion[item] else "")
    tree.bind("<Double-1>", toggle_selection)
    tree.bind("<Return>", toggle_selection)

    # Botones enviar y regresar
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=8)

    def enviar_seleccionados():
        usuarios_sel = [item_to_user[iid] for iid, sel in seleccion.items() if sel]
        if not usuarios_sel:
            messagebox.showwarning("Sin seleccionados", "No hay usuarios seleccionados. Selecciona con doble click sobre una fila.")
            return
        enviado = enviar_correos_lista(usuarios_sel)
        # Después del envío (o si el usuario canceló), preguntar si desea nueva búsqueda o volver
        if enviado:
            if messagebox.askyesno("Nueva búsqueda", "¿Deseas hacer una nueva consulta con otro filtro de días?"):
                # re-pedir días y refrescar
                new_dias = simpledialog.askinteger("Filtro de días", "¿Cuántos días antes del vencimiento deseas mostrar?", minvalue=1, maxvalue=180, parent=win)
                if not new_dias:
                    new_dias = dias
                # recalcular filtered
                new_filtered = [u for u in all_users if 0 <= u["dias"] <= new_dias]
                nonlocal_vars = {"filtered": new_filtered}
                # update filtered in closure:
                filtered.clear()
                filtered.extend(new_filtered)
                insertar_datos()
            else:
                # volver al menu principal
                win.destroy()
                parent_win.deiconify()
        else:
            # si no se envió (cancelado o error), preguntar si desea intentar otra búsqueda
            if messagebox.askyesno("Continuar", "¿Deseas hacer otra búsqueda?"):
                new_dias = simpledialog.askinteger("Filtro de días", "¿Cuántos días antes del vencimiento deseas mostrar?", minvalue=1, maxvalue=180, parent=win)
                if not new_dias:
                    new_dias = dias
                new_filtered = [u for u in all_users if 0 <= u["dias"] <= new_dias]
                filtered.clear()
                filtered.extend(new_filtered)
                insertar_datos()
            else:
                win.destroy()
                parent_win.deiconify()

    def regresar():
        win.destroy()
        parent_win.deiconify()

    ttk.Button(btn_frame, text="Enviar correo a seleccionados", command=enviar_seleccionados).pack(side="left", padx=8)
    ttk.Button(btn_frame, text="Regresar", command=regresar).pack(side="right", padx=8)

# ---- Ventana: Dashboard (gráfico circular) ----
def abrir_dashboard(parent_win, conn):
    parent_win.withdraw()
    win = tk.Toplevel()
    win.title("Dashboard de contraseñas")
    centrar_ventana(win, 800, 520)
    win.protocol("WM_DELETE_WINDOW", lambda: on_close_subwindow(win, parent_win))
    frame = ttk.Frame(win, padding=10)
    frame.pack(fill="both", expand=True)

    # Calcular datos
    all_users = consultar_usuarios(conn)
    expirados = [u for u in all_users if u["dias"] <= 0]
    proximos = [u for u in all_users if 1 <= u["dias"] <= 10]
    buenos = [u for u in all_users if u["dias"] > 20]

    counts = [len(expirados), len(proximos), len(buenos)]
    labels = [f"Expirados (≤0): {counts[0]}", f"Próximos (1-10): {counts[1]}", f"Bien (>20): {counts[2]}"]
    colors = None  # dejar matplotlib elegir

    # crear figura pie
    fig = Figure(figsize=(5,4), dpi=100)
    ax = fig.add_subplot(111)
    total = sum(counts)
    if total == 0:
        ax.text(0.5, 0.5, "No hay datos para mostrar", horizontalalignment='center', verticalalignment='center', fontsize=12)
    else:
        wedges, texts, autotexts = ax.pie(counts, autopct=lambda pct: f"{int(round(pct*total/100.0))} ({pct:.1f}%)", startangle=90)
        ax.legend(wedges, labels, title="Estados", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        ax.set_title("Estado de contraseñas (resumen)")

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

    # botones
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=8)
    ttk.Button(btn_frame, text="Regresar", command=lambda: (win.destroy(), parent_win.deiconify())).pack(side="right", padx=8)

# Cierre controlado de sub-ventanas: preguntar y volver al padre
def on_close_subwindow(win, parent_win):
    if messagebox.askyesno("Confirmar", "¿Deseas volver al menú principal?"):
        try:
            win.destroy()
        finally:
            parent_win.deiconify()
    else:
        return

# ==============================
# PUNTO DE ENTRADA
# ==============================
if __name__ == "__main__":
    ventana_login()