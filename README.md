# Checador Corporativo (Prototipo)

## Correccion aplicada (TemplateNotFound: login.html)

El error ocurria porque Flask no encontraba la carpeta "templates" al
ejecutarse con gunicorn en Render. Se corrigio forzando rutas ABSOLUTAS
explicitas para templates, static y la base de datos, basadas en la
ubicacion real de app.py:

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
    STATIC_DIR = os.path.join(BASE_DIR, "static")
    app = Flask(__name__, template_folder=TEMPLATES_DIR, static_folder=STATIC_DIR)

Ademas se agrego un log de diagnostico al iniciar la app que escribe en los
logs de Render si la carpeta templates existe o no, y que archivos contiene.

## MUY IMPORTANTE: verifica en GitHub que la carpeta templates SI se subio

1. Entra a tu repositorio en GitHub.
2. Debes poder navegar dentro de una carpeta llamada "templates" y ver:
   login.html, checador.html, admin_login.html, admin_panel.html,
   admin_empleados.html.
3. Debes poder navegar dentro de una carpeta llamada "static" y ver:
   style.css.
4. Si no ves esas carpetas o estan vacias, vuelve a subirlas arrastrando
   la carpeta completa (con sus archivos dentro) directamente en la pagina
   de tu repositorio en GitHub.

## Como desplegarlo en Render (gratis)

1. Crea una cuenta en render.com.
2. Sube esta carpeta completa a un repositorio de GitHub, con app.py,
   requirements.txt, runtime.txt, Procfile, la carpeta templates/ y la
   carpeta static/ en la RAIZ del repositorio.
3. En Render: "New +" -> "Web Service" -> conecta tu repositorio.
4. Configura (o deja que se autodetecte via render.yaml):
   - Build Command: pip install -r requirements.txt
   - Start Command: gunicorn app:app
5. Si tu repo tiene los archivos dentro de una subcarpeta, especifica esa
   carpeta en "Root Directory" dentro de Settings.
6. Revisa siempre la pestaña "Logs" del servicio en Render.

IMPORTANTE antes de compartir la URL con tu equipo:
- Cambia app.secret_key y ADMIN_PASSWORD dentro de app.py.

## Como correrlo en tu propia computadora (pruebas locales)

1. Instala Python 3.11.
2. Dentro de la carpeta "checador_app": pip install -r requirements.txt
3. Ejecuta: python app.py
4. Abre: http://localhost:5000

## Carga de empleados

- empleados_pins.csv: los 23 empleados con su PIN asignado (1001-1023).
- plantilla_carga_empleados.csv: mismo listado con columnas vacias para
  llenar el horario de cada persona.
- cargar_empleados.py: script para dar de alta/actualizar automaticamente
  a todos los empleados en la base de datos.

## Notas de seguridad

- Cambia app.secret_key y ADMIN_PASSWORD antes de usarlo con datos reales.
- Usa siempre HTTPS si se aloja en internet publico.
- Respalda periodicamente el archivo checador.db.
