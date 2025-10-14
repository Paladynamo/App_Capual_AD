# 🔐 AD Password Expiry Notifier

Aplicación de escritorio en **Python (Tkinter)** desarrollada por **Eduardo "PaladynamoX" Lizama C.**  
Permite consultar, visualizar y notificar por correo electrónico a los usuarios de **Active Directory (AD)** cuya contraseña está próxima a expirar.

---

## 🧩 Características principales

- Conexión segura a un **servidor LDAP/Active Directory**.
- Interfaz gráfica intuitiva creada con **Tkinter**.
- Selección del usuario autenticado mediante un **combobox de agentes**.
- Consulta de usuarios activos cuya contraseña expira en un rango configurable de días.
- Visualización de los resultados en una **tabla ordenable**.
- Envío automático de **correos de aviso personalizados** a cada usuario afectado.
- Posibilidad de **actualizar el filtro de días** sin reiniciar la app.

---

⚙️ Requisitos del sistema:

-	Python 3.8+
-	Conexión a un servidor Active Directory (LDAP/LDAPS) accesible.
-	Cuenta con permisos de lectura sobre los atributos:
-	sAMAccountName
-	displayName
-	mail
-	msDS-UserPasswordExpiryTimeComputed
-	department

---

🧰 Dependencias
Instala las librerías necesarias con:
- pip install ldap3 <br>
(Tkinter y smtplib vienen incluidos en la instalación estándar de Python.)

---

🧠 Detalles técnicos

-	Lenguaje: Python 3
-	Interfaz: Tkinter + ttk
-	Conexión: ldap3 (LDAP sobre SSL/TLS)
-	Envío de correos: smtplib + MIMEText
-	Gestión de fechas: datetime
-	Autor: Eduardo “PaladynamoX” Lizama C.
-	Versión: 1.0.0 (2025)

---

📨 Contacto

Creado por Eduardo “PaladynamoX” Lizama C.
💼 Cooperativa Capual - Departamento de Soporte TI
📧 Contacto: eduardo.1994.arte@gmail.com

---

⭐ Si este proyecto te resultó útil, no olvides dejar una estrella en el repositorio.
