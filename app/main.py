import json
from pathlib import Path
from uuid import uuid4
from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from .database import Base, SessionLocal, engine
from .models import Meerwerk
from .schemas import MeerwerkInput
from .services import UPLOADS, create_excel, normalize_photo, send_email

Base.metadata.create_all(bind=engine)
app = FastAPI(title="Jan Bos Meerwerk")
app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

@app.get("/")
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/meerwerk")
async def submit(data_json: str = Form(...), fotos: list[UploadFile] = File(default=[]), db: Session = Depends(get_db)):
    photo_paths: list[Path] = []
    try:
        data = MeerwerkInput.model_validate_json(data_json)
        if len(fotos) > 12:
            raise HTTPException(status_code=400, detail="Maximaal 12 foto's toegestaan")
        for upload in fotos:
            if not upload.filename:
                continue
            if not (upload.content_type or "").startswith("image/"):
                raise HTTPException(status_code=400, detail=f"{upload.filename} is geen afbeelding")
            raw = await upload.read()
            if len(raw) > 12 * 1024 * 1024:
                raise HTTPException(status_code=400, detail=f"{upload.filename} is groter dan 12 MB")
            destination = UPLOADS / f"{uuid4().hex}.jpg"
            normalize_photo(raw, destination)
            photo_paths.append(destination)
        excel_path = create_excel(data, photo_paths)
        record = Meerwerk(
            opdrachtgever=data.opdrachtgever, opdrachtnummer=data.opdrachtnummer,
            object=data.object, datum=data.datum, medewerker=data.medewerker,
            omschrijving=data.omschrijving, reden=data.reden, uren=str(data.uren),
            materialen=data.materialen, kostenindicatie=data.kostenindicatie,
            foto_paden_json=json.dumps([str(p) for p in photo_paths]), excel_pad=str(excel_path),
        )
        db.add(record); db.commit(); db.refresh(record)
        send_email(excel_path, data, photo_paths)
        return {"ok": True, "id": record.id, "bestandsnaam": excel_path.name}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Verzenden mislukt: {exc}") from exc

@app.get("/api/meerwerk/{item_id}/download")
def download(item_id: int, db: Session = Depends(get_db)):
    record = db.get(Meerwerk, item_id)
    if not record or not Path(record.excel_pad).exists():
        raise HTTPException(status_code=404, detail="Meerwerkrapport niet gevonden")
    return FileResponse(record.excel_pad, filename=Path(record.excel_pad).name)
