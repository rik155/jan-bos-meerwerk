from datetime import datetime
from sqlalchemy import Column, DateTime, Integer, String, Text
from .database import Base

class Meerwerk(Base):
    __tablename__ = "meerwerk"
    id = Column(Integer, primary_key=True)
    opdrachtgever = Column(String(200), nullable=False)
    opdrachtnummer = Column(String(100), default="")
    object = Column(String(250), nullable=False)
    datum = Column(String(30), nullable=False)
    medewerker = Column(String(150), nullable=False)
    omschrijving = Column(Text, nullable=False)
    reden = Column(Text, default="")
    uren = Column(String(30), default="0")
    materialen = Column(Text, default="")
    kostenindicatie = Column(String(50), default="")
    foto_paden_json = Column(Text, default="[]")
    excel_pad = Column(Text, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
