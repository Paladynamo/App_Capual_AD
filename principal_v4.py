# principal_v3.py
# Versión 3 - Dashboard interactivo con detalle por categoría
# Integrado en un solo archivo según el proyecto del usuario.

from ldap3 import Server, Connection, ALL, SUBTREE
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog, filedialog
import sys
import matplotlib
# usar backend TkAgg para interacción en la UI
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os, sys

# ==============================
# CONFIGURACIÓN (ajusta según tu entorno)
# ==============================
AD_SERVER = 'ldaps://SRV_DC01_NEW.capual.cl'
BASE_DN = 'DC=capual,DC=cl'

# Limitar la búsqueda solo a estas OUs bajo OU=Capual
ALLOWED_OUS = [
    "OU=Areas de Apoyo,OU=Capual,DC=capual,DC=cl",
    "OU=Directorio,OU=Capual,DC=capual,DC=cl",
    "OU=Gerencia,OU=Capual,DC=capual,DC=cl",
    "OU=Oficinas,OU=Capual,DC=capual,DC=cl",
]


# Las credenciales de envío se solicitarán en tiempo de ejecución (no se almacenan en el código)
SMTP_REMITENTE = None
SMTP_PASSWORD = None
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Cache opcional de credenciales solo para la sesión actual (si el usuario lo permite)
_SMTP_CACHE = {"remitente": None, "password": None}

APP_CREDITOS = """App creada por Eduardo 'PaladynamoX' Lizama C.
Versión 4.0.0 - Año 2025"""

# Detecta si se está ejecutando desde .exe o desde .py
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

IMG_PATH = os.path.join(BASE_DIR, "img_teclas.png")
LOGO_PATH = os.path.join(BASE_DIR, "logo_capual_antiguo.png")
FAREWELL_LOGO_PATH = os.path.join(BASE_DIR, "kuriboh_logo_despedida.png")


# ==============================
# UTILIDADES UI
# ==============================
def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


# Estilo visual unificado (ligero) para toda la app
def setup_style(root):
    """Aplica un estilo visual con tonos verde metalizado y corrige
    el mapeo de colores para evitar textos invisibles en botones."""
    try:
        style = ttk.Style(root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Paleta (verde claro metalizado)
        base_bg = "#eef6f1"   # fondo muy claro verdoso
        border = "#b6d6c3"

        # Fondo general
        style.configure("TFrame", background=base_bg)
        style.configure("TLabelframe", background=base_bg)
        style.configure("TLabel", background=base_bg, foreground="#20302a", font=("Segoe UI", 10))

        # Botones: volvemos a un estilo seguro (sin colorear fondo) para que SIEMPRE se
        # renderice el texto correctamente en Windows.
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 6), foreground="#222")
        style.map(
            "TButton",
            background=[("active", "#e6f3ec")],  # leve tinte verde en hover
            foreground=[
                ("disabled", "#777"),
                ("active", "#222"),
                ("!disabled", "#222"),
            ],
        )

        # Tabla
        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=24,
            background="#ffffff",
            fieldbackground="#ffffff",
            bordercolor=border,
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI Semibold", 10),
            background="#d8efe2",
            foreground="#1f2a26",
            bordercolor=border,
        )
    except Exception:
        pass


# Ajusta la altura de la tabla y la ventana según el número de registros
def auto_ajustar_altura(*args, **kwargs):
    # Desactivado por solicitud: mantenemos firma por compatibilidad
    return


# ==============================
# UTILIDAD: Exportar Treeview a Excel con estilo, logo y resumen
# ==============================
def _ensure_package_silent(pkg: str, import_name: str = None) -> bool:
    """Intenta importar un paquete y si falla, lo instala en silencio con pip.
    Devuelve True si el paquete está disponible al finalizar."""
    import importlib, subprocess, sys
    name = import_name or pkg
    try:
        importlib.import_module(name)
        return True
    except Exception:
        pass

    # Evitar intentos de instalación en ejecutables congelados (.exe)
    if getattr(sys, 'frozen', False):
        return False

    cmds = [
        [sys.executable, "-m", "pip", "install", pkg, "--quiet", "--disable-pip-version-check"],
        [sys.executable, "-m", "pip", "install", pkg, "--user", "--quiet", "--disable-pip-version-check"],
    ]
    for cmd in cmds:
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            import importlib
            importlib.invalidate_caches()
            importlib.import_module(name)
            return True
        except Exception:
            continue
    return False


def _ensure_excel_deps(parent_window=None) -> bool:
    """Asegura 'openpyxl' y 'Pillow' (para insertar imágenes)."""
    ok_openpyxl = _ensure_package_silent("openpyxl")
    ok_pillow = _ensure_package_silent("Pillow", import_name="PIL")
    if not ok_openpyxl:
        try:
            messagebox.showerror(
                "Dependencia faltante",
                "No fue posible cargar o instalar 'openpyxl'.\n\n" \
                "Si estás ejecutando un .exe, incluye 'openpyxl' en el empaquetado."
            )
        except Exception:
            pass
        return False
    # Pillow es opcional (solo para logo). Si falla, seguimos sin imagen.
    return True


def export_tree_to_excel(parent_window, tree: ttk.Treeview, titulo: str = "Reporte"):
    # Intentar importar dependencias; si faltan, instalarlas silenciosamente cuando sea posible.
    if not _ensure_excel_deps(parent_window):
        return

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule
    try:
        from openpyxl.drawing.image import Image as XLImage
    except Exception:
        XLImage = None

    # Elegir ruta
    path = filedialog.asksaveasfilename(
        parent=parent_window,
        title="Guardar como",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        initialfile=f"{titulo.replace(' ', '_')}.xlsx",
    )
    if not path:
        return

    # Preparar datos desde el Treeview
    cols = list(tree["columns"]) or []
    include_idx = [i for i, c in enumerate(cols) if str(c).strip().lower() != "sel"]
    headers = [cols[i] for i in include_idx]
    rows = [tree.item(iid, 'values') for iid in tree.get_children()]
    rows = [[r[i] for i in include_idx] for r in rows]

    # Calcular resumen por categorías a partir de "Días restantes"
    def _parse_int_safe(v):
        try:
            return int(str(v).strip())
        except Exception:
            return None

    dias_col_idx = None
    for idx, h in enumerate(headers):
        if "día" in str(h).lower():
            dias_col_idx = idx
            break
    counts = {"Bien (16-90)": 0, "Próximos (1-15)": 0, "Expirados (<=0)": 0, "Sin dato": 0}
    if dias_col_idx is not None:
        for r in rows:
            v = _parse_int_safe(r[dias_col_idx])
            if v is None:
                counts["Sin dato"] += 1
            elif v <= 0:
                counts["Expirados (<=0)"] += 1
            elif 1 <= v <= 15:
                counts["Próximos (1-15)"] += 1
            elif 16 <= v <= 90:
                counts["Bien (16-90)"] += 1
            else:
                # Fuera de rango superior, no contamos o podríamos sumarlo en Bien
                counts["Bien (16-90)"] += 0
    else:
        # No hay columna de días; solo creamos hoja de datos
        pass

    wb = openpyxl.Workbook()
    # Hoja de resumen primero
    ws_res = wb.active
    ws_res.title = "Resumen"
    ws = wb.create_sheet("Datos")

    # Estilos base
    verde = "27ae60"  # sin # para fills
    verde_claro = "d8efe2"
    gris_claro = "f7fbf9"
    thin = Side(style="thin", color="b6d6c3")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Logo (si es posible) para ambas hojas
    top_row = 1
    start_col_for_title = 2
    try:
        if os.path.exists(LOGO_PATH) and 'XLImage' in locals() and XLImage is not None:
            img1 = XLImage(LOGO_PATH)
            try:
                img1.width, img1.height = 140, int(140 * 0.45)
            except Exception:
                pass
            ws.add_image(img1, "A1")
            img2 = XLImage(LOGO_PATH)
            try:
                img2.width, img2.height = 140, int(140 * 0.45)
            except Exception:
                pass
            ws_res.add_image(img2, "A1")
            start_col_for_title = 2
    except Exception:
        pass

    last_col = max(len(headers), 1)
    # Título (fila 1)
    ws.merge_cells(start_row=top_row, start_column=start_col_for_title, end_row=top_row, end_column=last_col+1)
    c = ws.cell(row=top_row, column=start_col_for_title, value=titulo)
    c.font = Font(size=16, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=verde)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # Subtítulo (fecha) fila 2
    ws.merge_cells(start_row=top_row+1, start_column=start_col_for_title, end_row=top_row+1, end_column=last_col+1)
    c2 = ws.cell(row=top_row+1, column=start_col_for_title, value=f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    c2.font = Font(size=10, color="24543b")
    c2.alignment = Alignment(horizontal="center")

    # Encabezados (fila 4)
    header_row = top_row + 3
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=j, value=h)
        cell.font = Font(bold=True, color="1f2a26")
        cell.fill = PatternFill("solid", fgColor=verde_claro)
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde

    # Datos
    data_start = header_row + 1
    for i, r in enumerate(rows, start=data_start):
        for j, v in enumerate(r, start=1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.border = borde
            # Alinear ciertas columnas
            header = headers[j-1].lower()
            if "día" in header:
                cell.alignment = Alignment(horizontal="center")
            elif "fecha" in header:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
        # Zebra stripe
        if (i - data_start) % 2 == 0:
            for j in range(1, len(headers)+1):
                ws.cell(row=i, column=j).fill = PatternFill("solid", fgColor=gris_claro)

    # Auto ancho de columnas
    for j, h in enumerate(headers, start=1):
        col_letter = get_column_letter(j)
        max_len = len(str(h))
        for i in range(data_start, data_start + len(rows)):
            val = ws.cell(row=i, column=j).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 50)

    # Filtros, panes congelados
    ws.auto_filter.ref = f"A{header_row}:" + get_column_letter(len(headers)) + f"{header_row}"
    ws.freeze_panes = f"A{data_start}"

    # Formato condicional para "Días restantes" (<=0 rojo, 1-15 amarillo, 16-90 verde)
    try:
        dias_idx = next((idx+1 for idx, h in enumerate(headers) if "día" in str(h).lower()), None)
        if dias_idx and rows:
            col_letter = get_column_letter(dias_idx)
            rng = f"{col_letter}{data_start}:{col_letter}{data_start + len(rows) - 1}"
            rojo = PatternFill("solid", fgColor="f8d7da")
            amarillo = PatternFill("solid", fgColor="fff3cd")
            verde_fill = PatternFill("solid", fgColor="d4edda")

            ws.conditional_formatting.add(rng, CellIsRule(operator='lessThanOrEqual', formula=['0'], fill=rojo))
            ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['1','15'], fill=amarillo))
            ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['16','90'], fill=verde_fill))

            # Número entero para la columna de días
            for i in range(data_start, data_start + len(rows)):
                ws.cell(row=i, column=dias_idx).number_format = "0"
    except Exception:
        pass

    # Construir hoja Resumen con KPI y gráfico (antes de guardar)
    try:
        # Título y subtítulo
        last_col_res = 4
        ws_res.merge_cells(start_row=top_row, start_column=start_col_for_title, end_row=top_row, end_column=last_col_res)
        c = ws_res.cell(row=top_row, column=start_col_for_title, value=f"Resumen – {titulo}")
        c.font = Font(size=16, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=verde)
        c.alignment = Alignment(horizontal="center", vertical="center")

        ws_res.merge_cells(start_row=top_row+1, start_column=start_col_for_title, end_row=top_row+1, end_column=last_col_res)
        c2 = ws_res.cell(row=top_row+1, column=start_col_for_title, value=f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c2.font = Font(size=10, color="24543b")
        c2.alignment = Alignment(horizontal="center")

        table_start_row = top_row + 3
        ws_res.cell(row=table_start_row, column=2, value="Estado").font = Font(bold=True)
        ws_res.cell(row=table_start_row, column=3, value="Cantidad").font = Font(bold=True)

        estados = ["Bien (16-90)", "Próximos (1-15)", "Expirados (<=0)"]
        if counts.get("Sin dato", 0) > 0:
            estados.append("Sin dato")

        for i, est in enumerate(estados, start=1):
            ws_res.cell(row=table_start_row + i, column=2, value=est)
            ws_res.cell(row=table_start_row + i, column=3, value=counts.get(est, 0)).number_format = "0"

        # Estética de tabla
        for r in range(table_start_row, table_start_row + len(estados) + 1):
            for cidx in (2, 3):
                cell = ws_res.cell(row=r, column=cidx)
                cell.border = borde
                if r == table_start_row:
                    cell.fill = PatternFill("solid", fgColor=verde_claro)
                elif (r - table_start_row) % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=gris_claro)

        # Auto ancho
        ws_res.column_dimensions['B'].width = 25
        ws_res.column_dimensions['C'].width = 12

        # Gráfico de dona
        try:
            from openpyxl.chart import PieChart, Reference
            chart = PieChart()
            chart.title = "Distribución por estado"
            chart.holeSize = 50
            data = Reference(ws_res, min_col=3, min_row=table_start_row, max_row=table_start_row + len(estados))
            labels = Reference(ws_res, min_col=2, min_row=table_start_row + 1, max_row=table_start_row + len(estados))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
            # Colores a los puntos (verde, amarillo, rojo, gris)
            try:
                from openpyxl.chart.series import DataPoint
                colors = ["6aa84f", "f1c232", "e06666", "999999"]
                for idx in range(len(estados)):
                    dp = DataPoint(idx=idx)
                    dp.graphicalProperties.solidFill = colors[idx]
                    chart.series[0].data_points.append(dp)
            except Exception:
                pass
            ws_res.add_chart(chart, "E5")
        except Exception:
            pass
    except Exception:
        pass

    # Guardar al final, ya con Resumen
    try:
        wb.save(path)
        messagebox.showinfo("Exportar a Excel", "Archivo Excel exportado correctamente.", parent=parent_window)
    except Exception as e:
        messagebox.showerror("Exportar a Excel", f"No se pudo guardar el archivo:\n{e}", parent=parent_window)


def despedida_final(segundos: int = 5):
    """Muestra una ventana de despedida centrada, sin botones ni contador, con logo, y cierra al cabo de 'segundos'."""
    try:
        # Ocultar ventana raíz si existe para que solo se vea la despedida
        if root_all and root_all.winfo_exists():
            try:
                root_all.withdraw()
            except Exception:
                pass

        # Crear diálogo de despedida
        dlg = tk.Toplevel()
        setup_style(dlg)
        dlg.title("Despedida")
        dlg.resizable(False, False)
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)  # impedir cierre manual
        width, height = 520, 260
        dlg.transient(None)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)

        # Mensaje de créditos/firma
        msg = ttk.Label(
            frm,
            text=f"Gracias por usar esta aplicación.\n{APP_CREDITOS}",
            anchor="center",
            justify="center"
        )
        msg.pack(fill="x", pady=(0, 10))

        # Logo personal (escalado para caber en la ventana)
        logo_label = ttk.Label(frm)
        logo_label.pack(pady=(0, 4))
        try:
            from PIL import Image, ImageTk  # type: ignore
            if os.path.exists(FAREWELL_LOGO_PATH):
                with Image.open(FAREWELL_LOGO_PATH) as im:
                    max_w = width - 64  # margen dentro del cuadro
                    max_h = 110
                    im.thumbnail((max_w, max_h), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                    photo = ImageTk.PhotoImage(im)
                    logo_label.configure(image=photo)
                    # Guardar referencia para evitar GC
                    dlg._farewell_photo = photo
        except Exception:
            # Si PIL no está disponible o la imagen falla, continuar sin logo
            pass

        # Centrar después de construir contenido real
        try:
            dlg.update_idletasks()
            real_w = dlg.winfo_width() or width
            real_h = dlg.winfo_height() or height
            centrar_ventana(dlg, int(real_w), int(real_h))
            # Recentrar tras un pequeño retraso por si el sistema ajusta bordes/título
            dlg.after(120, lambda: centrar_ventana(dlg, dlg.winfo_width(), dlg.winfo_height()))
        except Exception:
            centrar_ventana(dlg, width, height)

        def finalizar():
            try:
                dlg.destroy()
            except Exception:
                pass
            try:
                if root_all and root_all.winfo_exists():
                    root_all.destroy()
            except Exception:
                pass
            sys.exit(0)

        # Programar cierre automático sin mostrar contador
        dlg.after(max(1000, int(segundos) * 1000), finalizar)

        # Asegurar que quede al frente y centrado
        try:
            dlg.lift()
            dlg.attributes('-topmost', True)
            dlg.after(200, lambda: dlg.attributes('-topmost', False))
            # Recentrar otra vez por seguridad tras el cambio de z-order/topmost
            dlg.after(220, lambda: centrar_ventana(dlg, dlg.winfo_width(), dlg.winfo_height()))
        except Exception:
            pass
    except Exception:
        # Cierre inmediato en caso de imprevistos
        try:
            if root_all and root_all.winfo_exists():
                root_all.destroy()
        except Exception:
            pass
        sys.exit(0)


def confirmar_y_cerrar(ventana):
    if messagebox.askyesno("Confirmar salida", "¿Deseas salir del programa?"):
        despedida_final(5)
    else:
        return

# ==============================
# FUNCIONES LDAP / CONSULTAS
# ==============================
def conectar_ldap(username, password):
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
    try:
        if not msds_value or int(msds_value) <= 0:
            return None
        expiry_date = datetime.fromtimestamp(int(msds_value) / 1e7 - 11644473600)
        return expiry_date
    except Exception:
        return None


def consultar_usuarios(conn):
    """
    Consulta los usuarios válidos desde el AD, restringido a las OUs indicadas.
    - Excluye usuarios sin correo.
    - Excluye usuarios con descripción: Auxiliar, Vigilante Privado o Guardia.
    - Solo busca en las subcarpetas válidas de 'Capual'.
    """

    filter_query = (
        "(&"
        "(objectCategory=person)"
        "(objectClass=user)"
        "(!(userAccountControl:1.2.840.113556.1.4.803:=2))"        # no deshabilitados
        "(!(userAccountControl:1.2.840.113556.1.4.803:=65536))"     # no 'password never expires'
        "(!(sAMAccountName=*$))"                                    # no cuentas de servicio
        "(!(sAMAccountName=Administrador))"
        "(mail=*)"                                                  # ✅ solo usuarios con correo
        ")"
    )

    attributes = [
        "sAMAccountName",
        "displayName",
        "mail",
        "msDS-UserPasswordExpiryTimeComputed",
        "department",
        "description",
        "distinguishedName",
    ]

    # 🔒 Solo en estas unidades organizativas
    ALLOWED_OUS = [
        "OU=Areas de Apoyo,OU=Capual,DC=capual,DC=cl",
        "OU=Directorio,OU=Capual,DC=capual,DC=cl",
        "OU=Gerencia,OU=Capual,DC=capual,DC=cl",
        "OU=Oficinas,OU=Capual,DC=capual,DC=cl",
    ]

    # 🚫 Descripciones que deben excluirse del resultado
    EXCLUDED_DESC = ["auxiliar", "vigilante privado", "guardia"]

    results = []
    now = datetime.now()

    for base_ou in ALLOWED_OUS:
        try:
            conn.search(base_ou, filter_query, SUBTREE, attributes=attributes)
        except Exception:
            continue

        for entry in conn.entries:
            try:
                sAM = str(entry["sAMAccountName"]) if entry["sAMAccountName"].value else ""
                if not sAM:
                    continue

                # Excluir si descripción coincide con alguno de los términos prohibidos
                desc = str(entry["description"]) if entry["description"].value else ""
                if any(ex in desc.lower() for ex in EXCLUDED_DESC):
                    continue

                # Validar correo (doble seguridad)
                mail = str(entry["mail"]) if entry["mail"].value else ""
                if not mail or "@" not in mail:
                    continue

                display = str(entry["displayName"]) if entry["displayName"].value else sAM
                dept = str(entry["department"]) if entry["department"].value else ""
                expiry_raw = entry["msDS-UserPasswordExpiryTimeComputed"].value
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
                    "expira": expiry_dt.strftime("%d/%m/%Y %H:%M"),
                    "descripcion": desc
                })
            except Exception:
                continue

    return results




# ==============================
# ENVÍO DE CORREOS (HTML + imagen embebida)
# ==============================
from email.mime.image import MIMEImage

def enviar_correos_lista(usuarios):
    if not usuarios:
        messagebox.showwarning("Aviso", "No hay usuarios seleccionados para enviar correos.")
        return False


def pedir_credenciales_smtp(parent):
    """Ventana modal para pedir correo remitente y contraseña.
    Devuelve (remitente, password) o (None, None) si se cancela.
    Añade opción 'Recordar en esta sesión'."""
    remit_var = tk.StringVar(value=_SMTP_CACHE.get("remitente") or "")
    pass_var = tk.StringVar(value="")
    remember_var = tk.BooleanVar(value=False)

    dlg = tk.Toplevel(parent)
    setup_style(dlg)
    dlg.title("Credenciales SMTP")
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.grab_set()
    centrar_ventana(dlg, 460, 200)
    dlg.protocol("WM_DELETE_WINDOW", lambda: (remit_var.set(""), pass_var.set(""), dlg.destroy()))

    frm = ttk.Frame(dlg, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Correo remitente (From):").grid(row=0, column=0, sticky="w", pady=(4,2))
    e_user = ttk.Entry(frm, textvariable=remit_var, width=40)
    e_user.grid(row=1, column=0, sticky="we")
    e_user.focus()

    ttk.Label(frm, text="Contraseña:").grid(row=2, column=0, sticky="w", pady=(8,2))
    e_pass = ttk.Entry(frm, textvariable=pass_var, show="*", width=40)
    e_pass.grid(row=3, column=0, sticky="we")

    chk = ttk.Checkbutton(frm, text="Recordar durante esta sesión", variable=remember_var)
    chk.grid(row=4, column=0, sticky="w", pady=(8,4))

    btns = ttk.Frame(frm)
    btns.grid(row=5, column=0, sticky="e", pady=(8,0))

    result = {"ok": False}

    def aceptar():
        user = remit_var.get().strip()
        pw = pass_var.get()
        if not user or "@" not in user:
            messagebox.showwarning("Dato requerido", "Ingresa un correo remitente válido.", parent=dlg)
            return
        if not pw:
            messagebox.showwarning("Dato requerido", "Ingresa la contraseña del remitente.", parent=dlg)
            return
        if remember_var.get():
            _SMTP_CACHE["remitente"] = user
            _SMTP_CACHE["password"] = pw
        result["ok"] = True
        dlg.destroy()

    def cancelar():
        result["ok"] = False
        dlg.destroy()

    ttk.Button(btns, text="Cancelar", command=cancelar).pack(side="right", padx=(6,0))
    ttk.Button(btns, text="Continuar", command=aceptar).pack(side="right")

    e_pass.bind("<Return>", lambda e: aceptar())
    dlg.bind("<Escape>", lambda e: cancelar())

    dlg.wait_window()
    if result["ok"]:
        return remit_var.get().strip(), pass_var.get()
    return None, None


def enviar_correos_con_progreso(usuarios, parent):
    """Envía correos mostrando una barra de progreso con opción de cancelar.
    Devuelve True si se envió al menos un correo.
    """
    if not usuarios:
        messagebox.showwarning("Aviso", "No hay usuarios seleccionados para enviar correos.", parent=parent)
        return False

    # Pedir credenciales (From y password); permitir recordar en sesión
    remitente, password = pedir_credenciales_smtp(parent)
    if not remitente or not password:
        return False

    # Confirmación luego de ingresar credenciales
    if not messagebox.askyesno("Confirmar envío", f"¿Desea enviar correos a {len(usuarios)} usuarios?\n\nDesde: {remitente}", parent=parent):
        return False

    progress_win = tk.Toplevel(parent)
    setup_style(progress_win)
    progress_win.title("Enviando correos…")
    progress_win.transient(parent)
    progress_win.grab_set()
    centrar_ventana(progress_win, 420, 140)

    frm = ttk.Frame(progress_win, padding=12)
    frm.pack(fill="both", expand=True)

    lbl = ttk.Label(frm, text="Preparando…")
    lbl.pack(fill="x", pady=(0,8))
    pbar = ttk.Progressbar(frm, mode="determinate", maximum=len(usuarios))
    pbar.pack(fill="x")

    cancel = tk.BooleanVar(value=False)
    def do_cancel():
        cancel.set(True)
    ttk.Button(frm, text="Cancelar", command=do_cancel).pack(pady=(10,0))

    parent.update_idletasks()

    enviados = 0
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(remitente, password)

        for idx, u in enumerate(usuarios, start=1):
            if cancel.get():
                break
            correo = u.get("correo")
            if not correo or "@" not in correo:
                continue

            msg = MIMEMultipart("related")
            msg["From"] = remitente
            msg["To"] = correo
            msg["Subject"] = "⚠️ Aviso: Tu contraseña está próxima a expirar"

            html_body = f"""
            <html>
            <body style="font-family:Segoe UI, sans-serif; color:#333;">
                <p>Estimado/a <b>{u.get('nombre','')}</b>,</p>
                <p>Tu contraseña expira en <b>{u.get('dias','-')} días</b> (el {u.get('expira','-')}).<br>
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
                <p>🏢 Departamento: {u.get('departamento','No especificado')}<br>
                👤 Usuario: {u.get('usuario','')}</p>
                <p>Saludos cordiales,<br>
                <b>Departamento de Soporte TI</b><br>
                Capual - Cooperativa de Ahorro y Crédito</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(html_body, "html"))
            try:
                with open(IMG_PATH, "rb") as f:
                    img = MIMEImage(f.read())
                    img.add_header("Content-ID", "<img_teclas>")
                    img.add_header("Content-Disposition", "inline", filename="img_teclas.png")
                    msg.attach(img)
            except FileNotFoundError:
                pass

            try:
                server.send_message(msg)
                enviados += 1
            except Exception:
                pass

            lbl.config(text=f"Enviando {idx}/{len(usuarios)}…")
            pbar['value'] = idx
            progress_win.update_idletasks()

        server.quit()
    except Exception as e:
        messagebox.showerror("Error envío", f"No se pudieron enviar los correos:\n{e}", parent=progress_win)
    finally:
        progress_win.destroy()

    if enviados > 0:
        messagebox.showinfo("Envío finalizado", f"Correos enviados: {enviados}", parent=parent)
        return True
    else:
        return False


# ==============================
# VENTANAS / FLUJO
# ==============================
root_all = None

def ventana_login():
    global root_all
    login_win = tk.Tk()
    root_all = login_win
    login_win.title(f"Iniciar sesión - {APP_CREDITOS}")
    setup_style(login_win)
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
    usuario_entry.focus()

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
            messagebox.showerror("Error", "Credenciales inválidas o no se pudo conectar al AD. Intenta nuevamente.")
            usuario_entry.focus()
            status_lbl.config(text="")

    ttk.Button(btn_frame, text="Iniciar sesión", command=intentar_login).grid(row=0, column=0, padx=6)
    ttk.Button(btn_frame, text="Salir", command=lambda: confirmar_y_cerrar(login_win)).grid(row=0, column=1, padx=6)

    login_win.mainloop()

# -----------------------------
# Ventana principal / Menú
# -----------------------------
def ventana_principal(conn):
    global root_all
    main_win = tk.Tk()
    root_all = main_win
    main_win.title(f"Menú Principal - {APP_CREDITOS}")
    setup_style(main_win)
    centrar_ventana(main_win, 560, 300)
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

    # Nuevo botón: búsqueda global por nombre/usuario/correo
    btn3 = ttk.Button(frm, text="🔎 Buscar por nombre o correo", width=36,
                      command=lambda: abrir_busqueda_usuario(main_win, conn))
    btn3.pack(pady=8)

    ttk.Label(frm, text=APP_CREDITOS, font=("Segoe UI", 9)).pack(side="bottom", pady=(12,0))

    main_win.mainloop()

# -----------------------------
# Usuarios próximos a expirar
# (se mantiene funcionalidad previa con orden/filtrado implementados en otra función)
# -----------------------------
def abrir_usuarios_proximos(parent_win, conn):
    parent_win.withdraw()
    win = tk.Toplevel()
    win.title("Usuarios próximos a expirar")
    setup_style(win)
    centrar_ventana(win, 980, 540)
    try:
        win.minsize(900, 520)
    except Exception:
        pass
    win.protocol("WM_DELETE_WINDOW", lambda: on_close_subwindow(win, parent_win))
    frame = ttk.Frame(win, padding=10)
    frame.pack(fill="both", expand=True)

    dias = simpledialog.askinteger("Filtro de días", "¿Cuántos días antes del vencimiento deseas mostrar?", minvalue=1, maxvalue=180, parent=win)
    if not dias:
        dias = 10

    all_users = consultar_usuarios(conn)
    filtered = [u for u in all_users if 0 <= u["dias"] <= dias]

    cols = ("Sel", "Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
    # Reducimos altura inicial para dejar espacio a los botones siempre visibles
    tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="none", height=15)
    for c in cols:
        tree.heading(c, text=c)
        if c == "Sel":
            tree.column(c, width=60, anchor="center")
        elif c == "Días restantes":
            tree.column(c, width=110, anchor="center")
        else:
            tree.column(c, width=150, anchor="w")
    tree.pack(side="top", fill="both", expand=True)

    seleccion = {}
    item_to_user = {}

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

    def toggle_selection(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        seleccion[item] = not seleccion[item]
        tree.set(item, "Sel", "✓" if seleccion[item] else "")
    tree.bind("<Double-1>", toggle_selection)
    tree.bind("<Return>", toggle_selection)

    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=8)

    # Exportadores (CSV/Excel)
    def exportar_csv():
        filas = [tree.item(i, 'values') for i in tree.get_children()]
        if not filas:
            messagebox.showinfo("Exportar", "No hay datos para exportar.", parent=win)
            return
        path = filedialog.asksaveasfilename(parent=win, defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path:
            return
        import csv
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            writer.writerows(filas)
        messagebox.showinfo("Exportar", "Archivo CSV exportado.", parent=win)

    def exportar_excel():
        export_tree_to_excel(win, tree, titulo=f"Usuarios próximos ({dias} días)")

    def enviar_seleccionados():
        usuarios_sel = [item_to_user[iid] for iid, sel in seleccion.items() if sel]
        if not usuarios_sel:
            messagebox.showwarning("Sin seleccionados", "No hay usuarios seleccionados. Selecciona con doble click sobre una fila.")
            return
        enviado = enviar_correos_con_progreso(usuarios_sel, win)
        if enviado:
            if messagebox.askyesno("Nueva búsqueda", "¿Deseas hacer una nueva consulta con otro filtro de días?"):
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
        else:
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

    # Zona izquierda: exportaciones
    exp_frame = ttk.Frame(btn_frame)
    exp_frame.pack(side="left")
    ttk.Button(exp_frame, text="Exportar Excel", command=exportar_excel).pack(side="left", padx=6)
    ttk.Button(exp_frame, text="Exportar CSV", command=exportar_csv).pack(side="left", padx=6)

    ttk.Button(btn_frame, text="Enviar correo a seleccionados", command=enviar_seleccionados).pack(side="left", padx=8)
    ttk.Button(btn_frame, text="Regresar", command=regresar).pack(side="right", padx=8)

# -----------------------------
# Dashboard mejorado e interactivo
# -----------------------------
def abrir_dashboard(parent_win, conn):
    parent_win.withdraw()
    win = tk.Toplevel()
    win.title("Dashboard de contraseñas")
    setup_style(win)
    centrar_ventana(win, 900, 560)
    win.protocol("WM_DELETE_WINDOW", lambda: on_close_subwindow(win, parent_win))
    frame = ttk.Frame(win, padding=10)
    frame.pack(fill="both", expand=True)

    # Obtener y clasificar usuarios según nueva política:
    # - Bien: dias >= 16 y <= 90
    # - Próximos: dias <= 15 and dias > 0
    # - Expirados: dias <= 0
    all_users = consultar_usuarios(conn)
    bien = [u for u in all_users if u["dias"] >= 16 and u["dias"] <= 90]
    proximos = [u for u in all_users if 1 <= u["dias"] <= 15]
    expirados = [u for u in all_users if u["dias"] <= 0]

    counts = [len(bien), len(proximos), len(expirados)]
    labels = [f"Bien (16-90): {counts[0]}", f"Próximos (1-15): {counts[1]}", f"Expirados (<=0): {counts[2]}"]
    colors = ["#6aa84f", "#f1c232", "#e06666"]  # verde, amarillo, rojo

    fig = Figure(figsize=(6,4), dpi=100)
    ax = fig.add_subplot(111)

    total = sum(counts)
    if total == 0:
        ax.text(0.5, 0.5, "No hay datos para mostrar", horizontalalignment='center', verticalalignment='center', fontsize=12)
    else:
        wedges, texts, autotexts = ax.pie(
            counts,
            labels=None,
            colors=colors,
            startangle=90,
            wedgeprops=dict(width=0.5),
            autopct="%1.1f%%"
        )

        ax.axis('equal')
        ax.set_title("Estado de contraseñas (resumen)")
        ax.legend(wedges, labels, title="Estados", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        fig.tight_layout()  # <<< 🔥 Ajuste automático


    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    widget = canvas.get_tk_widget()
    widget.pack(side="top", fill="both", expand=True)

    # etiquetas con conteos
    cont_frame = ttk.Frame(frame)
    cont_frame.pack(fill="x", pady=(6,8))
    ttk.Label(cont_frame, text=f"🟢 Bien (16-90 días): {counts[0]}").pack(side="left", padx=8)
    ttk.Label(cont_frame, text=f"🟡 Próximos (1-15 días): {counts[1]}").pack(side="left", padx=8)
    ttk.Label(cont_frame, text=f"🔴 Expirados (≤0 días): {counts[2]}").pack(side="left", padx=8)

    # Función que abre una ventana con los usuarios de la categoría
    def abrir_ventana_categoria(nombre_cat, usuarios_cat):
        modal = tk.Toplevel()
        modal.title(f"{nombre_cat} - Usuarios")
        setup_style(modal)
        centrar_ventana(modal, 980, 520)
        modal.transient(win)
        modal.grab_set()

        frm = ttk.Frame(modal, padding=10)
        frm.pack(fill="both", expand=True)

        # Buscador
        buscador_frame = ttk.Frame(frm)
        buscador_frame.pack(fill="x", pady=(0,8))
        ttk.Label(buscador_frame, text="🔍 Buscar:").pack(side="left", padx=(0,6))
        entrada_busq = ttk.Entry(buscador_frame, width=40)
        entrada_busq.pack(side="left", padx=(0,6))

        # Agregamos columna de selección visual "Sel" como primera columna
        cols = ("Sel", "Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
        # Reducimos la altura para reservar espacio a la botonera inferior
        tree = ttk.Treeview(frm, columns=cols, show="headings", height=15, selectmode="none")
        for c in cols:
            tree.heading(c, text=c)
            if c == "Sel":
                tree.column(c, width=60, anchor="center")
            elif c == "Días restantes":
                tree.column(c, width=120, anchor="center")
            else:
                tree.column(c, width=140, anchor="w")
        tree.pack(fill="both", expand=True)

        seleccion = {}
        item_to_user = {}

        # Insertar datos
        def insertar_tabla(datos):
            for r in tree.get_children():
                tree.delete(r)
            for u in datos:
                iid = tree.insert("", "end", values=("", u["usuario"], u["nombre"], u["correo"],
                                                     u["departamento"], u.get("dias","-"), u.get("expira","-")))
                seleccion[iid] = False
                item_to_user[iid] = u

        insertar_tabla(usuarios_cat)

        # Selección con doble clic o Enter
        def toggle_selection(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            seleccion[item] = not seleccion[item]
            tree.set(item, "Sel", "✓" if seleccion[item] else "")
            tree.item(item, tags=("selected",) if seleccion[item] else ())
            tree.tag_configure("selected", background="#cce5ff")

        tree.bind("<Double-1>", toggle_selection)
        tree.bind("<Return>", toggle_selection)

        # Ordenar columnas
        def ordenar_tabla(col, reverse=False):
            datos = [(tree.set(k, col), k) for k in tree.get_children('')]
            try:
                datos.sort(key=lambda t: (int(t[0]) if str(t[0]).isdigit() else str(t[0]).lower()), reverse=reverse)
            except Exception:
                datos.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
            for idx, (val, k) in enumerate(datos):
                tree.move(k, '', idx)
            tree.heading(col, command=lambda: ordenar_tabla(col, not reverse))

        for c in cols:
            tree.heading(c, text=c, command=lambda c=c: ordenar_tabla(c))

        # Filtro
        def filtrar():
            texto = entrada_busq.get().strip().lower()
            for item in tree.get_children():
                valores = tree.item(item, "values")
                visible = any(texto in str(v).lower() for v in valores)
                tree.detach(item) if not visible else tree.reattach(item, '', 'end')

        def limpiar():
            entrada_busq.delete(0, tk.END)
            insertar_tabla(usuarios_cat)

        ttk.Button(buscador_frame, text="Buscar", command=filtrar).pack(side="left", padx=(0,6))
        ttk.Button(buscador_frame, text="Limpiar", command=limpiar).pack(side="left")
        # accesos rápidos
        entrada_busq.bind("<Return>", lambda e: filtrar())
        modal.bind("<Escape>", lambda e: (entrada_busq.delete(0, tk.END), insertar_tabla(usuarios_cat)))

        # Botones inferiores
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="Regresar", command=lambda: modal.destroy()).pack(side="right", padx=6)

        # Exportar a Excel la vista actual
        def exportar_excel_local():
            export_tree_to_excel(modal, tree, titulo=f"{nombre_cat} - Usuarios")
        ttk.Button(btns, text="Exportar Excel", command=exportar_excel_local).pack(side="left", padx=6)

        # Solo para “Próximos” y “Expirados”
        if nombre_cat in ("Próximos", "Expirados"):
            def enviar_todos():
                visibles = []
                for item in tree.get_children():
                    usuario_key = tree.item(item, "values")[1]  # columna Usuario (después de Sel)
                    for u in usuarios_cat:
                        if u["usuario"] == usuario_key:
                            visibles.append(u)
                            break
                if not visibles:
                    messagebox.showwarning("Sin destinatarios", "No hay usuarios visibles para enviar.")
                    return
                if messagebox.askyesno("Confirmar envío", f"¿Enviar correos a {len(visibles)} usuarios visibles?"):
                    enviar_correos_con_progreso(visibles, modal)

            def enviar_seleccionados():
                seleccionados = [item_to_user[i] for i, sel in seleccion.items() if sel]
                if not seleccionados:
                    messagebox.showwarning("Sin seleccionados", "No has seleccionado ningún usuario (doble clic para marcar).")
                    return
                if messagebox.askyesno("Confirmar envío", f"¿Enviar correos a {len(seleccionados)} usuarios seleccionados?"):
                    enviar_correos_con_progreso(seleccionados, modal)

            ttk.Button(btns, text="Enviar correos a todos", command=enviar_todos).pack(side="left", padx=6)
            ttk.Button(btns, text="Enviar correos seleccionados", command=enviar_seleccionados).pack(side="left", padx=6)



    # conectar eventos de click sobre wedges
    if total > 0:
        # hacer que cada wedge sea "pickable"
        for w in (ax.patches if hasattr(ax, 'patches') else []):
            pass

        # Usar pick_event: asignar picker a wedges devueltos por pie
        # Cuando creamos el pie devolvimos 'wedges' arriba
        try:
            for w in wedges:
                w.set_picker(True)
        except Exception:
            pass

        def on_pick(event):
            artist = event.artist
            # identificar índice
            try:
                idx = wedges.index(artist)
            except Exception:
                return
            if idx == 0:
                abrir_ventana_categoria("Bien", bien)
            elif idx == 1:
                abrir_ventana_categoria("Próximos", proximos)
            elif idx == 2:
                abrir_ventana_categoria("Expirados", expirados)

        canvas.mpl_connect('pick_event', on_pick)

        # tooltip al mover el ratón
        tooltip = tk.Label(frame, text="", background="#ffffe0", relief="solid", bd=1)
        tooltip.place_forget()

        def on_move(event):
            if not event.inaxes:
                tooltip.place_forget()
                return
            found = False
            for i, w in enumerate(wedges):
                contains, _ = w.contains(event)
                if contains:
                    # posicion absoluta del cursor en ventana
                    x, y = event.guiEvent.x_root, event.guiEvent.y_root
                    tooltip.config(text=labels[i])
                    # colocar tooltip cerca del cursor (coordenadas relativas a la ventana principal)
                    tooltip.place(x=event.x + 20, y=event.y + 20)
                    found = True
                    break
            if not found:
                tooltip.place_forget()

        canvas.mpl_connect('motion_notify_event', on_move)

    # botones
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=8)
    ttk.Button(btn_frame, text="Regresar", command=lambda: (win.destroy(), parent_win.deiconify())).pack(side="right", padx=8)

# -----------------------------
# Cierre modal
# -----------------------------
def on_close_subwindow(win, parent_win):
    if messagebox.askyesno("Confirmar", "¿Deseas volver al menú principal?"):
        try:
            win.destroy()
        finally:
            parent_win.deiconify()
    else:
        return

# -----------------------------
# Utilidad para escapar filtros LDAP
# -----------------------------
def escape_ldap_filter_value(value: str) -> str:
    """Escapa caracteres especiales para filtros LDAP (RFC4515)."""
    repl = {
        "\\": r"\\5c",
        "*": r"\\2a",
        "(": r"\\28",
        ")": r"\\29",
        "\x00": r"\\00",
    }
    out = []
    for ch in value:
        out.append(repl.get(ch, ch))
    return "".join(out)

# -----------------------------
# Búsqueda global por nombre/usuario/correo en todo el AD
# -----------------------------
def buscar_usuarios_global(conn, termino: str, *, incluir_deshabilitados: bool = False, incluir_pwd_never_expires: bool = True, base_dn: str = None):
    termino = (termino or "").strip()
    if not termino:
        return []

    term = escape_ldap_filter_value(termino)

    # Construir filtro de forma flexible
    filtros = [
        "(objectCategory=person)",
        "(objectClass=user)",
    ]
    if not incluir_deshabilitados:
        filtros.append("(!(userAccountControl:1.2.840.113556.1.4.803:=2))")
    if not incluir_pwd_never_expires:
        filtros.append("(!(userAccountControl:1.2.840.113556.1.4.803:=65536))")

    filtros.append(
        "(|" +
        f"(displayName=*{term}*)" +
        f"(sAMAccountName=*{term}*)" +
        f"(mail=*{term}*)" +
        ")"
    )

    filter_query = "(&" + "".join(filtros) + ")"

    attributes = [
        "sAMAccountName",
        "displayName",
        "mail",
        "msDS-UserPasswordExpiryTimeComputed",
        "department",
        "description",
        "distinguishedName",
    ]

    EXCLUDED_DESC = ["auxiliar", "vigilante privado", "guardia"]

    results = []
    now = datetime.now()
    try:
        conn.search(base_dn or BASE_DN, filter_query, SUBTREE, attributes=attributes)
    except Exception:
        return []

    for entry in conn.entries:
        try:
            sAM = str(entry["sAMAccountName"]) if entry["sAMAccountName"].value else ""
            if not sAM:
                continue

            desc = str(entry["description"]) if entry["description"].value else ""
            if any(ex in desc.lower() for ex in EXCLUDED_DESC):
                continue

            mail = str(entry["mail"]) if entry["mail"].value else ""
            display = str(entry["displayName"]) if entry["displayName"].value else sAM
            dept = str(entry["department"]) if entry["department"].value else ""
            dn = str(entry["distinguishedName"]) if entry["distinguishedName"].value else ""
            expiry_raw = entry["msDS-UserPasswordExpiryTimeComputed"].value
            expiry_dt = msds_to_datetime(expiry_raw)
            dias_restantes = (expiry_dt - now).days if expiry_dt else None

            results.append({
                "usuario": sAM,
                "nombre": display,
                "correo": mail,
                "departamento": dept,
                "dias": dias_restantes if dias_restantes is not None else "-",
                "expira": expiry_dt.strftime("%d/%m/%Y %H:%M") if expiry_dt else "-",
                "descripcion": desc,
                "dn": dn,
            })
        except Exception:
            continue

    return results

# -----------------------------
# Ventana de búsqueda global
# -----------------------------
def abrir_busqueda_usuario(parent_win, conn):
    parent_win.withdraw()
    win = tk.Toplevel()
    win.title("Buscar usuario por nombre, usuario o correo")
    setup_style(win)
    centrar_ventana(win, 980, 560)
    win.protocol("WM_DELETE_WINDOW", lambda: on_close_subwindow(win, parent_win))

    frame = ttk.Frame(win, padding=10)
    frame.pack(fill="both", expand=True)

    # Barra de búsqueda
    top = ttk.Frame(frame)
    top.pack(fill="x", pady=(0,8))
    ttk.Label(top, text="🔎 Término a buscar:").pack(side="left", padx=(0,6))
    entrada = ttk.Entry(top, width=34)
    entrada.pack(side="left")

    # Búsqueda simplificada (solo término)
    def realizar_busqueda():
        termino = entrada.get().strip()
        if not termino:
            messagebox.showwarning("Dato requerido", "Ingresa un nombre, usuario o correo a buscar.")
            return
        start = datetime.now()
        datos = buscar_usuarios_global(conn, termino)
        insertar_datos(datos)
        elapsed = (datetime.now() - start).total_seconds()
        status_lbl.config(text=f"Coincidencias: {len(datos)}   •   Tiempo: {elapsed:.2f}s")

    ttk.Button(top, text="Buscar", command=realizar_busqueda).pack(side="left", padx=8)
    ttk.Button(top, text="Limpiar", command=lambda: (entrada.delete(0, tk.END), insertar_datos([]), status_lbl.config(text=""))).pack(side="left")

    status_lbl = ttk.Label(frame, text="")
    status_lbl.pack(fill="x", pady=(0,6))

    cols = ("Sel", "Usuario", "Nombre", "Correo", "Departamento", "Días restantes", "Fecha de expiración")
    # Reducimos altura para garantizar visibilidad de la botonera inferior
    tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="none", height=15)
    for c in cols:
        tree.heading(c, text=c)
        if c == "Sel":
            tree.column(c, width=60, anchor="center")
        elif c in ("Días restantes", "Fecha de expiración"):
            tree.column(c, width=130, anchor="center")
        else:
            tree.column(c, width=160, anchor="w")
    tree.pack(side="top", fill="both", expand=True)

    seleccion = {}
    item_to_user = {}

    def insertar_datos(datos):
        for r in tree.get_children():
            tree.delete(r)
        seleccion.clear()
        item_to_user.clear()
        for u in datos:
            vals = ("", u["usuario"], u["nombre"], u["correo"], u["departamento"], str(u["dias"]), u["expira"])
            iid = tree.insert("", "end", values=vals)
            seleccion[iid] = False
            item_to_user[iid] = u
    # altura fija: sin auto-resize

    def toggle_selection(event):
        item = tree.identify_row(event.y)
        if not item:
            return
        seleccion[item] = not seleccion[item]
        tree.set(item, "Sel", "✓" if seleccion[item] else "")

    tree.bind("<Double-1>", toggle_selection)
    tree.bind("<Return>", toggle_selection)

    # Botonera
    btns = ttk.Frame(frame)
    btns.pack(fill="x", pady=8)

    def enviar_sel():
        usuarios_sel = [item_to_user[iid] for iid, sel in seleccion.items() if sel]
        if not usuarios_sel:
            messagebox.showwarning("Sin seleccionados", "No hay usuarios seleccionados. Doble clic para marcar.")
            return
        enviar_correos_con_progreso(usuarios_sel, win)

    def export_csv():
        rows = [tree.item(i, 'values') for i in tree.get_children()]
        if not rows:
            messagebox.showinfo("Exportar", "No hay datos para exportar.", parent=win)
            return
        path = filedialog.asksaveasfilename(parent=win, defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path:
            return
        import csv
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            writer.writerows(rows)
        messagebox.showinfo("Exportar", "Archivo CSV exportado.", parent=win)

    def copiar_portapapeles():
        rows = [tree.item(i, 'values') for i in tree.get_children()]
        if not rows:
            return
        text = "\n".join("\t".join(map(str, r)) for r in rows)
        try:
            win.clipboard_clear()
            win.clipboard_append(text)
        except Exception:
            pass

    ttk.Button(btns, text="Exportar CSV", command=export_csv).pack(side="left", padx=6)
    ttk.Button(btns, text="Exportar Excel", command=lambda: export_tree_to_excel(win, tree, titulo="Búsqueda de usuarios")).pack(side="left", padx=6)
    ttk.Button(btns, text="Copiar", command=copiar_portapapeles).pack(side="left", padx=6)
    ttk.Button(btns, text="Enviar correo a seleccionados", command=enviar_sel).pack(side="left", padx=6)
    ttk.Button(btns, text="Regresar", command=lambda: (win.destroy(), parent_win.deiconify())).pack(side="right", padx=6)

    # Atajos: Enter para buscar, Escape para limpiar
    entrada.bind("<Return>", lambda e: realizar_busqueda())
    win.bind("<Escape>", lambda e: (entrada.delete(0, tk.END), insertar_datos([]), status_lbl.config(text="")))

# ==============================
# PUNTO DE ENTRADA
# ==============================
if __name__ == "__main__":
    ventana_login()
