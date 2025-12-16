# Retornos - Web app

Este repo contiene una primera migración del Excel "Diario - Retorno Portafolios V1.1.xlsm" a una aplicación web en Python.

¿Qué se implementó?

- Un endpoint POST /retornos que acepta un CSV con el mismo formato de RetornosV21.csv (PowerQuery) y devuelve los registros transformados en JSON.
- Lógica de procesamiento en src/processing.py: salta 5 filas, promueve encabezados, detecta delimitador, convierte tipos y filtra por Account opcional.
- Script de arranque run.sh y requirements.txt para dependencias.

Cómo ejecutar (local):

1. Instalar dependencias:
   pip install -r requirements.txt
2. Ejecutar:
   ./run.sh
   el servidor se expondrá en http://0.0.0.0:12000

Ejemplo de uso (curl):

curl -F "file=@RetornosV21.csv" http://127.0.0.1:12000/retornos

Siguientes pasos:

- Implementar los cálculos adicionales que hace el Excel/VBA (TWRR agregados, resúmenes, generación de PDF/imagenes y envío por correo).
- Crear UI web para subir archivos y mostrar resultados.
- Añadir pruebas unitarias para validar las transformaciones y cálculos.