"""Configuración pública de DCO Feedback.

Solo usa una clave publicable/anon. El acceso a datos queda controlado por RLS.
NUNCA agregues una clave sb_secret_ ni service_role en aplicaciones distribuidas.
"""
import os

SUPABASE_URL = os.getenv(
    "DCO_FEEDBACK_URL", "https://fsljiwjzmjunrxlwklft.supabase.co"
).rstrip("/")

SUPABASE_PUBLIC_KEY = os.getenv(
    "DCO_FEEDBACK_PUBLIC_KEY", "sb_publishable_knGkQFd5lz6szELcW8XICA_V2jofeN6"
).strip()

POLL_INTERVAL_MS = 5 * 60 * 1000


def configurado():
    return (
        SUPABASE_URL.startswith("https://")
        and "TU-PROYECTO" not in SUPABASE_URL
        and not SUPABASE_URL.endswith("/rest/v1")
        and SUPABASE_PUBLIC_KEY not in ("", "PEGAR_CLAVE_PUBLICABLE")
        and (
            SUPABASE_PUBLIC_KEY.startswith("sb_publishable_")
            or SUPABASE_PUBLIC_KEY.startswith("eyJ")  # anon JWT legado
        )
    )
