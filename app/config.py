from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    brevo_api_key: str = ""
    mail_from: str = "info@janbosschilderwerken.nl"
    mail_to: str = "info@janbosschilderwerken.nl"
    mail_from_name: str = "Jan Bos Meerwerk"
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

settings = Settings()
