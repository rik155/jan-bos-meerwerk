# Jan Bos Meerwerk

Losse FastAPI-app voor het digitaal vastleggen van meerwerk.

## Functies
- Opdrachtgever, opdrachtnummer, object, datum en medewerker
- Omschrijving, reden, extra uren, materialen en kostenindicatie
- Maximaal 12 foto's vanaf camera of galerij
- Handtekening met vinger of stylus
- Automatisch Excel-rapport met foto's en handtekening
- Excel en foto's via Brevo naar kantoor
- SQLite archief en download-endpoint
- Geschikt voor Render via Docker

## Render
Maak een nieuwe Web Service met Docker en stel in:
- `BREVO_API_KEY`
- `MAIL_FROM`
- `MAIL_TO`
- `MAIL_FROM_NAME`

Voor blijvende opslag op Render is een persistent disk op `/app/data` aanbevolen.
