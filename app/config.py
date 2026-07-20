import os

class Settings:
    @property
    def brevo_api_key(self) -> str:
        return os.getenv('BREVO_API_KEY', '').strip()

    @property
    def mail_to(self) -> str:
        return (os.getenv('MAIL_TO') or os.getenv('DEFAULT_EXPORT_EMAIL') or '').strip()

    @property
    def mail_from(self) -> str:
        return (os.getenv('MAIL_FROM') or os.getenv('SMTP_USERNAME') or os.getenv('DEFAULT_EXPORT_EMAIL') or '').strip()

    mail_from_name = 'Jan Bos Meerwerk'

settings = Settings()
