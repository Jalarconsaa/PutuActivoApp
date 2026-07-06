# PUTÚ ACTIVO — Sistema de Gestión v5.0

## Instalación

1. Descomprime esta carpeta
2. Pon `datosgym.xlsx` dentro (para importación inicial)
3. Doble clic en `INICIAR_PUTU_ACTIVO.bat`

El `.bat` instala todo automáticamente:
```
streamlit · pandas · plotly · openpyxl · qrcode · pillow
```

## Accesos

| Usuario | Contraseña | Rol |
|---|---|---|
| admin | putu2025 | Administrador |
| entrenador | gym123 | Entrenador |
| recepcion | recep1 | Recepción |

Sesión dura 5 horas sin necesidad de volver a ingresar clave.

## Primera vez — Importar datos

1. Entra al menú **⚙️ Base de Datos**
2. Tab **📥 Importar**
3. Selecciona modo `datosgym.xlsx original`
4. Sube el archivo y presiona **Importar**

## Estructura de carpetas generada automáticamente

```
putu_activo/
├── app.py
├── INICIAR_PUTU_ACTIVO.bat
├── putu_activo.db          ← base de datos SQLite (se crea sola)
├── fotos_clientes/         ← fotos subidas desde la ficha
└── rutinas_pdf/            ← PDFs de rutinas por cliente
```

## Módulos v5.0

| Módulo | Novedades |
|---|---|
| 🏠 Dashboard | Donut activos por sexo, vencimientos con WhatsApp |
| 👥 Clientes | Todos los campos editables, foto, rutina PDF, QR, historial pagos, evaluaciones en ficha, contrato y derecho a saber completos con texto negro, enlace directo a Pagos |
| ➕ Nuevo Cliente | Fecha vencimiento automática según período |
| 💳 Pagos | Pre-llena el cliente desde la ficha, tipo de plan + frecuencia + medio de pago, WhatsApp confirmación |
| ✅ Asistencia | Teclado numérico en pantalla grande para kiosco/tablet, display aeropuerto con clientes en sala y próximas clases |
| 🏃 Clases & Talleres | Pase diario incluido, título, tipo adultos/niños/taller |
| 💰 Ingresos & Egresos | Resumen mensual, egresos manuales, ventas productos |
| 📊 Reportes | Activos, inactivos, vencimientos urgentes con WhatsApp, cumpleaños con WhatsApp, flujo de caja mensual |
| ⚙️ Base de Datos | Una sola DB SQLite · Exportar TODO en un Excel · Importar Excel unificado o datosgym original |

## Red local

```
http://<tu-IP-local>:8501
```
