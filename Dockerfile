# Imagen para correr BradentonApp en cualquier ambiente Linux (Render/Toolbox
# incluido) sin depender de que alguien instale nada a mano en el servidor.
#
# Resuelve el pendiente de "Tesseract necesita instalarse a nivel de sistema
# operativo" (ver CLAUDE.md, sección "Portabilidad a Render/Linux"): con este
# Dockerfile, Tesseract se instala solo cada vez que se construye la imagen
# -- no hace falta que Jorge (ni nadie) lo instale a mano en Render.
#
# No cambia nada del comportamiento en la PC del usuario (Windows) -- este
# archivo solo se usa si Render (u otro ambiente) construye la app a partir
# de esta imagen; localmente se sigue corriendo igual que siempre.
FROM python:3.12-slim

# tesseract-ocr: el motor de OCR real que pytesseract necesita para leer
# PDFs escaneados/fotografiados (Reporte Diario, Lottery, Gettel Pagos,
# Proveedores, y los controles mensuales que reusan ese mismo motor) --
# ver ocr_utils.py. Sin esto, esos módulos fallan con un ImportError claro
# (no un crash), pero simplemente no pueden usarse.
# tzdata: la zona horaria de Florida (abajo, TZ). Sin ella, en Render "hoy"
# es la fecha UTC y después de las 20:00 de Florida la app ya está en el día
# (o el mes) siguiente (J6 de AUDITORIA_2026-10.md).
RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr tzdata \
    && rm -rf /var/lib/apt/lists/*

ENV TZ=America/New_York
# Los print() de la app salen enseguida en los logs de Render (sin buffer).
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Copiar solo requirements.txt primero para que Docker cachee esta capa --
# reconstruir la imagen tras un cambio de código no reinstala todas las
# dependencias de Python de nuevo, solo cuando requirements.txt cambia.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# webapp.py ya lee la variable PORT que Render asigna (ver el "if __name__
# == '__main__'" al final de ese archivo) y escucha en host="0.0.0.0" --
# no hace falta ningún argumento extra acá. Con DATABASE_URL, antes de
# atender aplica las migraciones de Alembic que falten (db.upgrade_schema).
CMD ["python", "webapp.py"]
