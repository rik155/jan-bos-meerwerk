# Jan Bos Meerwerk

Losse FastAPI-app voor digitale meerwerkrapporten.

## Functies

- Opdrachtgever, opdrachtnummer, werkadres, datum en medewerker
- Omschrijving, reden, extra uren, materialen en kostenindicatie
- Maximaal 12 foto's vanaf camera of galerij
- Handtekening met vinger of stylus
- Excel-rapport met foto's en handtekening
- Excel en losse foto's per e-mail via SMTP
- Klaar voor Docker, GitHub en Render

## Render Environment Variables

Gebruik exact dezelfde SMTP-instellingen als bij de werkende Glaszetter-app:

- `DEFAULT_EXPORT_EMAIL`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `SMTP_FROM`
- `SMTP_USE_TLS`

`BREVO_API_KEY`, `MAIL_FROM` en `MAIL_TO` zijn voor deze versie niet nodig.
