from urllib.parse import quote_plus

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    # SQL Server connection
    DB_SERVER: str = "localhost"
    DB_PORT: int = 1433
    DB_NAME: str = "mydb"
    DB_USER: str = "app_user"
    DB_PASSWORD: str = "AppUser@2026#Secure"

    # JWT
    JWT_SECRET_KEY: str = "AQWZSXEDCRFVTGBYHNUJMKL1234567890"
    JWT_ALGORITHM: str = "HS256"
    JWT_ACCESS_TOKEN_EXPIRE_MINUTES: int = 60

    @property
    def database_url(self) -> str:
        password = quote_plus(self.DB_PASSWORD)
        return (
            f"mssql+pymssql://{self.DB_USER}:{password}"
            f"@{self.DB_SERVER}:{self.DB_PORT}/{self.DB_NAME}"
        )

    class Config:
        env_file = ".env"


settings = Settings()
