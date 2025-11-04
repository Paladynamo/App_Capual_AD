<p align="center">
	<img src="icono_app_edu_original.png" alt="Portada - AD Password Expiry Notifier" width="180">
</p>

# 🔐 AD Password Expiry Notifier

Notificador de expiración de contraseñas de Active Directory (versión 5.0)

<p align="center">
	<a href="https://www.python.org/"><img alt="Python" src="https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white"></a>
	<img alt="GUI" src="https://img.shields.io/badge/GUI-Tkinter-5A9?logo=python&logoColor=white">
	<a href="https://ldap3.readthedocs.io/"><img alt="LDAP3" src="https://img.shields.io/badge/LDAP-ldap3-0052CC"></a>
	<a href="https://openpyxl.readthedocs.io/"><img alt="Excel" src="https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoft-excel&logoColor=white"></a>
	<a href="https://matplotlib.org/"><img alt="Matplotlib" src="https://img.shields.io/badge/Charts-matplotlib-013243"></a>
</p>

Aplicación de escritorio en Python que consulta Active Directory, muestra el estado de expiración de contraseñas y permite enviar avisos personalizados por correo. Incluye exportación a Excel con estilo corporativo y un panel “Dashboard” con gráfico tipo dona.

---

## Índice

- [Novedades v5.0](#novedades-v50)
- [Características](#características)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Configuración](#configuración)
- [Uso rápido](#uso-rápido)
- [Exportación a Excel](#exportación-a-excel)
- [Envío de correos](#envío-de-correos)
- [Dashboard](#dashboard)
- [Empaquetado (PyInstaller)](#empaquetado-pyinstaller)
- [Recursos/Assets](#recursosassets)
- [Solución de problemas](#solución-de-problemas)
- [Créditos](#créditos)

---

## Novedades v5.0

- Dashboard V2 con mayor interactividad:
	- Filtros vivos por estado (Bien/Próximos/Expirados) y por rango de días (-30 a 90).
	- Gráfico dona con drill‑down por clic y tooltips al pasar el mouse.
	- Histograma de “días restantes”.
	- Top 10 usuarios más urgentes (doble clic abre propiedades).
	- Acceso rápido “Abrir vista filtrada” y fila de botones fija al fondo.
- Exportación a Excel “Resumen” reforzada:
	- Logo con reescalado nítido (Pillow LANCZOS) y títulos reubicados.
	- KPIs tipo “cards”, gráfico de dona, porcentajes con barras de datos y enlace a “Datos”.
	- Tabla de Top departamentos y “Top 10 más urgentes”.
	- Nota explicativa y auto‑ajuste de anchos.
- Envío de correos flexible: selector de método (Outlook o SMTP) con opción “Enviar como”.
- Búsqueda global en AD por nombre, usuario o correo (amplia y rápida).
- Estilo visual consolidado (ttk/clam) con textos legibles.

---

## Características

- Conexión segura a Active Directory vía ldap3 (LDAP/LDAPS).
- UI de escritorio con Tkinter/ttk, tablas ordenables y filtros rápidos.
- Avisos por correo con HTML e imagen embebida (instrucciones Ctrl+Alt+Supr).
- Exportación a CSV y a Excel con formato profesional y logo de la institución.
- Panel “Dashboard” con gráfico tipo dona (matplotlib) y accesos a listas por categoría.
- Compatibilidad con empaquetado a .exe (PyInstaller).

---

## Requisitos

- Python 3.8 o superior
- Conectividad al dominio de Active Directory
- Usuario con permisos de lectura en atributos: sAMAccountName, displayName, mail, msDS-UserPasswordExpiryTimeComputed, department

Dependencias principales (instalación típica):
- ldap3
- matplotlib
- openpyxl (Excel)
- Pillow (opcional, para insertar imágenes en Excel y procesar logos)

Tkinter y smtplib vienen con Python por defecto.

---

## Instalación

1) Clona o descarga este repositorio.
2) Crea (opcional) y activa un entorno virtual.
3) Instala dependencias:

```powershell
# Windows PowerShell
pip install ldap3 matplotlib openpyxl Pillow
```

---

## Configuración

Ajusta los valores del archivo `principal_v4.py` según tu entorno:

- `AD_SERVER`: URL del DC (ej: `ldaps://SRV_DC01_NEW.capual.cl`)
- `BASE_DN`: DN base del dominio (ej: `DC=capual,DC=cl`)
- `ALLOWED_OUS`: lista de OUs donde se restringe la consulta de usuarios
- `SMTP_SERVER` / `SMTP_PORT`: servidor y puerto SMTP (por defecto Office 365)
- Rutas de imágenes (se detectan automáticamente en el directorio de la app):
	- `IMG_PATH` (imagen de instrucciones para el correo)
	- `LOGO_PATH` (logo para Excel)
	- `FAREWELL_LOGO_PATH` (logo de despedida)

> Nota: el remitente y su contraseña NO están en el código. Se solicitarán al enviar correos y, si lo decides, se recordarán únicamente durante la sesión actual.

---

## Uso rápido

- Ejecuta la aplicación:

```powershell
python .\principal_v4.py
```

- Inicia sesión con tu usuario de dominio.
- Desde el menú:
	- “Usuarios próximos a expirar”: consulta por rango de días, permite seleccionar destinatarios y enviar correos.
	- “Dashboard de contraseñas”: muestra resumen con gráfico y acceso al detalle por categoría.
	- “Buscar por nombre o correo”: búsqueda global flexible en todo el AD.

---

## Exportación a Excel

- La hoja “Datos” incluye: encabezados con estilo, zebra striping, bordes, auto-ancho de columnas, filtros y formato condicional para “Días restantes”.
- La hoja “Resumen” agrega KPIs por estado (Bien, Próximos, Expirados y, si aplica, Sin dato) y un gráfico de dona con colores coherentes.
- Si `LOGO_PATH` existe, el logo se inserta en ambas hojas.

---

## Envío de correos

- Al presionar “Enviar correo…”, se abrirá un diálogo pidiendo el correo remitente y su contraseña.
- Puedes marcar “Recordar durante esta sesión” para no reingresarlos nuevamente.
- Los mensajes se envían en HTML e incluyen (si existe) la imagen `img_teclas.png` embebida.
- El progreso del envío se muestra en una ventana con barra de avance y opción de cancelar.

---

## Dashboard

- Dona con tres categorías: Bien (16–90), Próximos (1–15) y Expirados (≤0), con drill‑down por clic y tooltips.
- Filtros vivos por estado y rango de días, KPIs de conteo, histograma de distribución.
- Top 10 urgentes con doble clic a propiedades y botón “Abrir vista filtrada”.
- Botonera fija inferior para regresar/exportar sin perderla por tamaño de ventana.

---

## Empaquetado (PyInstaller)

Se incluye `principal_v4.spec`. Puedes usarlo o ejecutar un comando equivalente. Asegúrate de:

- Incluir los módulos de `openpyxl` y (opcionalmente) `Pillow` si deseas insertar imágenes en Excel.
- Empaquetar los recursos/imagenes: `img_teclas.png`, `logo_capual_antiguo.png`, `kuriboh_logo_despedida.png`.
- Probar el envío SMTP desde el ejecutable (TLS 587) para verificar conectividad y credenciales.

> Si ejecutas un .exe, la instalación automática de dependencias no está disponible; debes incluirlas en el empaquetado.

---

## Recursos/Assets

- `img_teclas.png` → insertada en el correo como imagen en línea.
- `logo_capual_antiguo.png` → insertado en Excel (Datos/Resumen).
- `kuriboh_logo_despedida.png` → mostrado en la ventana de despedida.

Coloca estos archivos junto al ejecutable o al script principal.

---

## Solución de problemas

- “No se pudo conectar al AD”: confirma `AD_SERVER`, credenciales y conectividad/puerto.
- “No se pudo guardar el Excel”: verifica permisos en la carpeta destino o cierra el archivo si ya estaba abierto.
- “No se pudieron enviar los correos”: revisa `SMTP_SERVER/PORT`, credenciales del remitente y conectividad TLS/587.
- El texto de los botones no se ve: la app fuerza un estilo seguro de Tkinter/ttk (clam) para mantener la legibilidad.

---

## Créditos

- Autor: **Eduardo “PaladynamoX” Lizama C.** — GitHub: [@Paladynamo](https://github.com/Paladynamo)
- Organización: **Cooperativa Capual – Departamento de Soporte TI**
- Versión de la app: **5.0.0 (2025)**
- Contacto: **eduardo.1994.arte@gmail.com**

> ¿Te fue útil? ⭐ ¡Apoya el proyecto con una estrella!
