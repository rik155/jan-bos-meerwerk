from pydantic import BaseModel, Field

class MeerwerkInput(BaseModel):
    opdrachtgever: str = Field(min_length=1, max_length=200)
    opdrachtnummer: str = Field(default="", max_length=100)
    object: str = Field(min_length=1, max_length=250)
    datum: str = Field(min_length=1, max_length=30)
    medewerker: str = Field(min_length=1, max_length=150)
    omschrijving: str = Field(min_length=1, max_length=5000)
    reden: str = Field(default="", max_length=3000)
    uren: float = Field(default=0, ge=0, le=1000)
    materialen: str = Field(default="", max_length=5000)
    kostenindicatie: str = Field(default="", max_length=50)
    handtekening: str = Field(min_length=20)
