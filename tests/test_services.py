from pathlib import Path
from PIL import Image
from app.schemas import MeerwerkInput
from app.services import create_excel
import base64, io

def test_excel(tmp_path):
    img=Image.new('RGB',(200,120),'white'); b=io.BytesIO(); img.save(b,'PNG')
    sig='data:image/png;base64,'+base64.b64encode(b.getvalue()).decode()
    photo=tmp_path/'foto.jpg'; Image.new('RGB',(400,300),'gray').save(photo)
    data=MeerwerkInput(opdrachtgever='Test',object='Werkadres',datum='2026-07-20',medewerker='Jan',omschrijving='Extra schilderwerk',uren=2.5,handtekening=sig)
    path=create_excel(data,[photo])
    assert Path(path).exists()
    Path(path).unlink()
