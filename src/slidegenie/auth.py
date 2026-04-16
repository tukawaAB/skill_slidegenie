"""Gemini API authentication.

Supports two authentication methods:
1. API Key: Set GOOGLE_API_KEY environment variable (takes priority)
2. ADC: Application Default Credentials via `gcloud auth application-default login`

Environment variables can be set in a .env file at the project root.

Ported from: ppt-addin/backend/services/google_cloud_client.py
             ppt-addin/backend/services/make_pptx/make_pptx_service.py
"""
import os
from pathlib import Path

from google import genai


def _load_dotenv(env_path: Path) -> None:
    """Load .env file into os.environ (only sets vars not already set)."""
    if not env_path.is_file():
        return
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip()
            # Remove surrounding quotes
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
                value = value[1:-1]
            # Don't override existing env vars
            if key not in os.environ:
                os.environ[key] = value


# Load .env from project root (slidegenie/ directory)
_project_root = Path(__file__).resolve().parent.parent.parent
_load_dotenv(_project_root / ".env")

_genai_client: genai.Client | None = None


def get_genai_client() -> genai.Client:
    """Get or create a singleton Gemini API client.

    Priority:
    1. GOOGLE_API_KEY env var → direct API key authentication
    2. ADC (Application Default Credentials) → VertexAI authentication
    """
    global _genai_client
    if _genai_client is not None:
        return _genai_client

    api_key = os.getenv("GOOGLE_API_KEY")

    if api_key:
        _genai_client = genai.Client(api_key=api_key)
    else:
        # ADC authentication
        from google.auth import default as google_auth_default

        SCOPES = ["https://www.googleapis.com/auth/cloud-platform"]
        credentials, detected_project = google_auth_default(scopes=SCOPES)
        project = os.getenv("GOOGLE_CLOUD_PROJECT") or detected_project
        if not project:
            raise EnvironmentError(
                "Could not resolve project. Set GOOGLE_CLOUD_PROJECT or "
                "configure ADC with 'gcloud auth application-default login'"
            )
        _genai_client = genai.Client(
            vertexai=True,
            project=project,
            location="global",
            credentials=credentials,
        )

    return _genai_client
