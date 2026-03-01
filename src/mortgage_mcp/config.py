"""Application configuration via environment variables."""

import json
import os
import tempfile

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    google_cloud_project: str = ""
    google_cloud_location: str = "northamerica-northeast1"
    google_application_credentials_json: str = ""
    mcp_auth_token: str = ""
    port: int = 8000
    gemini_model: str = "gemini-2.5-flash"

    model_config = {"env_prefix": "", "case_sensitive": False}

    def setup_gcp_credentials(self) -> None:
        """Write JSON credentials to a temp file and set the env var.

        This handles Render deployments where the service account key
        is stored as a JSON string in an environment variable.
        """
        if self.google_application_credentials_json and not os.environ.get(
            "GOOGLE_APPLICATION_CREDENTIALS"
        ):
            creds = json.loads(self.google_application_credentials_json)
            tmp = tempfile.NamedTemporaryFile(
                mode="w", suffix=".json", delete=False
            )
            json.dump(creds, tmp)
            tmp.close()
            os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = tmp.name


settings = Settings()
