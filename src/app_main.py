from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import Optional
import uvicorn
from pathlib import Path
from src.processing import process_retornos_csv
from src.visuals import generate_informe_images
from src.pdfgen import generate_pdf
from fastapi.responses import StreamingResponse, JSONResponse


BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Retornos API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Static files and templates
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/")
async def index(request: Request):
    """Página principal con UI para subir archivos y visualizar resultados"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/retornos")
async def retornos(file: UploadFile = File(...), account: Optional[str] = Form(None)):
    """Procesa un CSV de retornos y devuelve las filas filtradas y transformadas.

    - file: CSV con formato similar a RetornosV21.csv (PowerQuery salteaba 5 filas y promovía encabezados)
    - account: (opcional) filtro para la columna 'Account'
    """
    if not file.filename.lower().endswith(('.csv', '.txt')):
        raise HTTPException(status_code=400, detail="Se requiere un archivo CSV")

    try:
        contents = await file.read()
        results = process_retornos_csv(contents, account=account)
        # processing returns dict with rows/count/informe and snapshots
        images = []
        try:
            images = generate_informe_images(results.get('informe', []))
        except Exception:
            images = []
        # snapshots image
        snapshot_images = []
        try:
            from visuals import generate_snapshots_image
            snapshot_images = generate_snapshots_image(results.get('snapshots', []))
        except Exception:
            snapshot_images = []
        images.extend(snapshot_images)
        return {**results, 'images': images, 'snapshot_images': snapshot_images}
    finally:
        await file.close()


@app.post('/retornos/pdf')
async def retornos_pdf(selected_date: str = Form(...), file: UploadFile = File(None)):
    """Genera un PDF con la tabla de comparación para la fecha seleccionada.

    - selected_date: fecha elegida en el selector
    - file: (opcional) archivo CSV si se desea re-subir; si no se envía se usa el archivo local
    """
    try:
        if file and file.filename:
            contents = await file.read()
            results = process_retornos_csv(contents)
            await file.close()
        else:
            local_csv = BASE_DIR.parent / 'RetornosV21_test.csv'
            if not local_csv.exists():
                local_csv = BASE_DIR.parent / 'RetornosV21.csv'
            if not local_csv.exists():
                return JSONResponse(status_code=404, content={"detail": "Archivo local no encontrado para generar PDF"})
            contents = local_csv.read_bytes()
            results = process_retornos_csv(contents)

        pdf_bytes = generate_pdf(results, selected_date)
        return StreamingResponse(iter([pdf_bytes]), media_type='application/pdf', headers={"Content-Disposition": f"attachment; filename=Informe_{selected_date}.pdf"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})


@app.get('/retornos/local')
async def retornos_local(account: Optional[str] = None):
    """Procesa el archivo RetornosV21_test.csv que está en el repositorio sin que el usuario lo suba.

    Esto es útil para prototipado: usa el archivo local en la raíz del repo.
    """
    local_csv = BASE_DIR.parent / 'RetornosV21_test.csv'
    if not local_csv.exists():
        # fallback: try RetornosV21.csv
        local_csv = BASE_DIR.parent / 'RetornosV21.csv'
    if not local_csv.exists():
        raise HTTPException(status_code=404, detail=f"Archivo local no encontrado: {local_csv}")

    try:
        contents = local_csv.read_bytes()
        results = process_retornos_csv(contents, account=account)
        images = []
        try:
            images = generate_informe_images(results.get('informe', []))
        except Exception:
            images = []
        # snapshots image
        snapshot_images = []
        try:
            from visuals import generate_snapshots_image
            snapshot_images = generate_snapshots_image(results.get('snapshots', []))
        except Exception:
            snapshot_images = []
        images.extend(snapshot_images)
        return {**results, 'images': images, 'snapshot_images': snapshot_images}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    uvicorn.run("src.app_main:app", host="0.0.0.0", port=12000, log_level="info")
