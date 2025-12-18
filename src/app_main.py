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
async def retornos(
    file: UploadFile = File(...),
    control_file: UploadFile = File(None),
    query_file: UploadFile = File(None),
    account: Optional[str] = Form(None)
):
    """Procesa un CSV de retornos y devuelve las filas filtradas y transformadas.

    - file: CSV con formato similar a RetornosV21.csv (PowerQuery salteaba 5 filas y promovía encabezados)
    - control_file: (opcional) CSV similar a ControlReporte1.csv usado en PowerQuery
    - query_file: (opcional) CSV con datos equivalentes al resultado de la consulta SQL (Query1)
    - account: (opcional) filtro para la columna 'Account'
    """
    if not file or not file.filename.lower().endswith(('.csv', '.txt')):
        raise HTTPException(status_code=400, detail="Se requiere un archivo CSV principal (Retornos)")

    support_files = {}
    try:
        contents = await file.read()
        # read optional support files
        if control_file and control_file.filename:
            try:
                support_files['control'] = await control_file.read()
            finally:
                await control_file.close()
        if query_file and query_file.filename:
            try:
                support_files['query'] = await query_file.read()
            finally:
                await query_file.close()

        results = process_retornos_csv(contents, account=account, support_files=support_files)
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
async def retornos_pdf(
    request: Request,
    selected_date: str = Form(...),
    file: UploadFile = File(None),
    control_file: UploadFile = File(None),
    query_file: UploadFile = File(None),
):
    """Genera un PDF con la tabla de comparación para la fecha seleccionada.

    This endpoint also supports generating an HTML navigable report instead of PDF
    by passing ?format=html (returns rendered HTML template).
    """
    fmt = request.query_params.get('format', 'pdf')
    try:
        support_files = {}
        if file and file.filename:
            contents = await file.read()
            await file.close()
        else:
            local_csv = BASE_DIR.parent / 'RetornosV21_test.csv'
            if not local_csv.exists():
                local_csv = BASE_DIR.parent / 'RetornosV21.csv'
            if not local_csv.exists():
                return JSONResponse(status_code=404, content={"detail": "Archivo local no encontrado para generar PDF/HTML"})
            contents = local_csv.read_bytes()

        # optional support files from upload
        if control_file and control_file.filename:
            support_files['control'] = await control_file.read()
            await control_file.close()
        else:
            # try repo-local fallback
            local_control = BASE_DIR.parent / 'ControlReporte1.csv'
            if local_control.exists():
                support_files['control'] = local_control.read_bytes()

        if query_file and query_file.filename:
            support_files['query'] = await query_file.read()
            await query_file.close()
        else:
            local_query = BASE_DIR.parent / 'Query1.csv'
            if local_query.exists():
                support_files['query'] = local_query.read_bytes()

        results = process_retornos_csv(contents, support_files=support_files)

        if fmt == 'html':
            # Render the navigable HTML report using templates/report.html
            context = {"request": request, 'selected_date': selected_date}
            context.update(results)
            # ensure images exist in results
            context['images'] = results.get('images', [])
            return templates.TemplateResponse('report.html', context)

        pdf_bytes = generate_pdf(results, selected_date)
        return StreamingResponse(iter([pdf_bytes]), media_type='application/pdf', headers={"Content-Disposition": f"attachment; filename=Informe_{selected_date}.pdf"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": str(e)})


@app.post('/retornos/email')
async def retornos_email(
    selected_date: str = Form(...),
    smtp_server: str = Form('smtp.office365.com'),
    smtp_port: int = Form(587),
    smtp_user: str = Form(...),
    smtp_pass: str = Form(...),
    to_email: str = Form(...),
    subject: Optional[str] = Form(None),
    body: Optional[str] = Form(None),
    file: UploadFile = File(None),
):
    """Genera el PDF y envía por SMTP usando credenciales proporcionadas en tiempo de ejecución.

    Nota: las credenciales NO se guardan en servidor.
    """
    import smtplib
    from email.message import EmailMessage

    try:
        # obtain results from uploaded file or local
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

        # compose email
        msg = EmailMessage()
        msg['Subject'] = subject or f'Informe Retornos {selected_date}'
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg.set_content(body or 'Adjunto se encuentra el informe de retornos.')

        msg.add_attachment(pdf_bytes, maintype='application', subtype='pdf', filename=f'Informe_{selected_date}.pdf')

        # connect to SMTP and send
        server = smtplib.SMTP(smtp_server, int(smtp_port), timeout=20)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()

        return JSONResponse(status_code=200, content={"detail": "Correo enviado"})
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
