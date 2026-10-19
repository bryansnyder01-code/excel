"""Runtime settings, read from environment variables (or a .env file)."""

from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_prefix="SHEETCRAFT_", extra="ignore")

    # LLM
    model: str = "claude-sonnet-5-5"
    effort: str = "medium"  # low | medium | high | xhigh | max
    max_tokens: int = 32000
    # How many times the engine sends validation errors back to the model for a fix.
    repair_attempts: int = 1

    # API
    cors_origins: list[str] = ["http://localhost:3000"]
    max_prompt_chars: int = 4000

    # AI usage limits: Claude requests per rolling 24 hours, per signed-in user or per client IP.
    ai_limit_enabled: bool = True
    ai_limit_anonymous: int = 30
    ai_limit_signed_in: int = 200
    ai_limit_pro: int = 2000  # signed-in users on the Pro plan
    # Behind a reverse proxy: take the client IP from the first X-Forwarded-For hop.
    trust_proxy: bool = False

    # Accounts and cloud sync: the SQLite database lives at <data_dir>/sheetcraft.db.
    data_dir: str = "data"
    session_days: int = 30  # sessions end after this many days without use

    # Password reset emails link here; Google sign-in needs an OAuth client ID.
    app_url: str = "http://localhost:3000"
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    smtp_from: str = ""
    google_client_id: str = ""

    # Live editing across several backend instances: a Redis URL (redis://[:password@]host:6379/0).
    # Empty keeps rooms in this process only.
    redis_url: str = ""

    # Billing (Stripe). Empty secret key or price = billing off.
    stripe_secret_key: str = ""
    stripe_webhook_secret: str = ""
    stripe_price_pro: str = ""

    # Serve the built frontend (a static export) from this folder at "/", so one process runs
    # the whole app. Used by the desktop app; empty for the usual two-service setup.
    static_dir: str = ""

    # Once-a-minute background jobs: live data refresh and scheduled email reports (off in tests).
    background_jobs: bool = True


@lru_cache
def get_settings() -> Settings:
    return Settings()
