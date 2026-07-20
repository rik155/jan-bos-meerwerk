import base64
import io
import re
from datetime import datetime
from pathlib import Path
from typing import Iterable
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps
from .config import settings
from .schemas import MeerwerkInput

GENERATED = Path("data/generated")
UPLOADS = Path("data/uploads")
GENERATED.mkdir(parents=True, exist_ok=True)
UPLOADS.mkdir(parents=True, exist_ok=True)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER = PatternFill("solid", fgColor="EAF3EA")
GREEN = "218C42"


def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value.strip())
    return value[:60] or "meerwerk"


def decode_signature(data_url: str) -> bytes:
    if "," not in data_url:
        raise ValueError("Ongeldige handtekening")
    return base64.b64decode(data_url.split(",", 1)[1])


def normalize_photo(raw: bytes, destination: Path) -> None:
    image = Image.open(io.BytesIO(raw))
    image = ImageOps.exif_transpose(image).convert("RGB")
    image.thumbnail((1800, 1800))
    image.save(destination, "JPEG", quality=86, optimize=True)


def create_excel(data: MeerwerkInput, photo_paths: list[Path]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Meerwerk"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_margins.left = .3
    ws.page_margins.right = .3
    for idx, width in enumerate([21, 18, 18, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    logo = Path("app/static/images/jan-bos-logo.png")
    if logo.exists():
        img = XLImage(str(logo)); img.width = 245; img.height = 100; ws.add_image(img, "A1")
    ws.merge_cells("E1:H2"); ws["E1"] = "Meerwerkrapport"
    ws["E1"].font = Font(size=23, bold=True, color=GREEN)
    ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("E3:H3"); ws["E3"] = "Jan Bos Glas en Schilderwerken"
    ws["E3"].alignment = Alignment(horizontal="center")
    ws.merge_cells("E4:H4"); ws["E4"] = "024 - 388 6392  |  info@janbosschilderwerken.nl"
    ws["E4"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 34; ws.row_dimensions[2].height = 34

    fields = [
        ("Opdrachtgever", data.opdrachtgever, "Datum", data.datum),
        ("Opdrachtnummer", data.opdrachtnummer, "Medewerker", data.medewerker),
        ("Object / werkadres", data.object, "Aantal uren", data.uren),
    ]
    row = 6
    for l1, v1, l2, v2 in fields:
        ws[f"A{row}"] = l1; ws[f"A{row}"].font = Font(bold=True)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4); ws[f"B{row}"] = v1
        ws[f"E{row}"] = l2; ws[f"E{row}"].font = Font(bold=True)
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=8); ws[f"F{row}"] = v2
        for col in range(1, 9): ws.cell(row, col).border = BORDER
        row += 1

    sections = [
        ("Omschrijving van het meerwerk", data.omschrijving, 5),
        ("Reden / oorzaak", data.reden or "-", 3),
        ("Gebruikte materialen", data.materialen or "-", 4),
        ("Kostenindicatie", data.kostenindicatie or "Niet ingevuld", 1),
    ]
    row += 1
    for title, text, height in sections:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        ws.cell(row, 1).value = title; ws.cell(row, 1).font = Font(bold=True, color="FFFFFF")
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=GREEN)
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row + height - 1, end_column=8)
        ws.cell(row, 1).value = text; ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
        for rr in range(row, row + height):
            for cc in range(1, 9): ws.cell(rr, cc).border = BORDER
        row += height + 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row, 1).value = f"Foto's ({len(photo_paths)})"
    ws.cell(row, 1).font = Font(bold=True, color="FFFFFF"); ws.cell(row, 1).fill = PatternFill("solid", fgColor=GREEN)
    row += 1
    if photo_paths:
        for idx, photo in enumerate(photo_paths, 1):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row, 1).value = f"Foto {idx}"
            ws.cell(row, 1).font = Font(bold=True); ws.cell(row, 1).fill = HEADER
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row + 11, end_column=8)
            for rr in range(row, row + 12):
                for cc in range(1, 9): ws.cell(rr, cc).border = BORDER
            img = XLImage(str(photo)); img.width = 520; img.height = 300; ws.add_image(img, f"B{row}")
            row += 13
    else:
        ws.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=8)
        ws.cell(row, 1).value = "Geen foto's toegevoegd"; ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
        row += 3

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row, 1).value = "Handtekening klant / opdrachtgever"
    ws.cell(row, 1).font = Font(bold=True, color="FFFFFF"); ws.cell(row, 1).fill = PatternFill("solid", fgColor=GREEN)
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 4, end_column=8)
    for rr in range(row, row + 5):
        for cc in range(1, 9): ws.cell(rr, cc).border = BORDER
    sig_path = GENERATED / f"sig_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.png"
    Image.open(io.BytesIO(decode_signature(data.handtekening))).convert("RGBA").save(sig_path)
    sig = XLImage(str(sig_path)); sig.width = 280; sig.height = 90; ws.add_image(sig, f"C{row}")

    filename = f"Meerwerk_{safe_name(data.opdrachtgever)}_{safe_name(data.datum)}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    path = GENERATED / filename
    wb.save(path)
    sig_path.unlink(missing_ok=True)
    return path


def send_email(path: Path, data: MeerwerkInput, photo_paths: Iterable[Path]) -> str:
    if not settings.brevo_api_key:
        raise RuntimeError('BREVO_API_KEY ontbreekt of is leeg in Render')
    if not settings.mail_to:
        raise RuntimeError('Ontvanger ontbreekt: stel DEFAULT_EXPORT_EMAIL of MAIL_TO in')
    if not settings.mail_from:
        raise RuntimeError('Afzender ontbreekt: stel SMTP_USERNAME of MAIL_FROM in')

    photos = list(photo_paths)
    attachments = [{"content": base64.b64encode(path.read_bytes()).decode("ascii"), "name": path.name}]
    for photo in photos:
        attachments.append({"content": base64.b64encode(photo.read_bytes()).decode("ascii"), "name": photo.name})
    payload = {
        "sender": {"name": settings.mail_from_name, "email": settings.mail_from},
        "to": [{"email": settings.mail_to}],
        "subject": f"Meerwerk - {data.opdrachtgever} - {data.object}",
        "htmlContent": (
            f"<p>Nieuw meerwerkrapport van <b>{data.medewerker}</b>.</p>"
            f"<p><b>Opdrachtgever:</b> {data.opdrachtgever}<br>"
            f"<b>Object:</b> {data.object}<br><b>Datum:</b> {data.datum}<br>"
            f"<b>Uren:</b> {data.uren:g}<br><b>Aantal foto's:</b> {len(photos)}</p>"
        ),
        "attachment": attachments,
    }
    print(f'Brevo: meerwerk verzenden van {settings.mail_from} naar {settings.mail_to}', flush=True)
    response = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"api-key": settings.brevo_api_key, "content-type": "application/json", "accept": "application/json"},
        json=payload, timeout=60,
    )
    if not response.ok:
        raise RuntimeError(f'Brevo fout {response.status_code}: {response.text[:500]}')
    result = response.json() if response.content else {}
    message_id = result.get('messageId', '')
    if not message_id:
        raise RuntimeError(f'Brevo gaf geen messageId terug: {response.text[:500]}')
    print(f'Brevo bevestigd. messageId={message_id}', flush=True)
    return message_id
