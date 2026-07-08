# CHECK LIST SS.HH DE LA CORPORACION - imagen para despliegue web
FROM python:3.12-slim

WORKDIR /app

RUN pip install --no-cache-dir openpyxl==3.1.5

COPY checklist_sshh.py .

# /datos es donde vive el Excel (montar un volumen aquí para no perder datos)
ENV CARPETA_DATOS=/datos \
    SIN_NAVEGADOR=1 \
    PUERTO=8740 \
    PYTHONUNBUFFERED=1

VOLUME /datos
EXPOSE 8740

CMD ["python", "checklist_sshh.py"]
