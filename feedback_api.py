"""Cliente HTTPS ligero para DCO Feedback usando Supabase Auth y PostgREST.

Los compañeros usan autenticación anónima de Supabase, que NO equivale a
compartir las credenciales del portal DCO. El administrador usa una cuenta
separada de Supabase Auth y permisos RLS en la base de datos.
"""
import json
import os
import threading
import time
from pathlib import Path
from urllib.parse import quote

import requests

from feedback_config import SUPABASE_URL, SUPABASE_PUBLIC_KEY, configurado


class FeedbackError(RuntimeError):
    pass


def _ruta_sesion():
    if os.name == "nt":
        base = Path(os.getenv("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif os.sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.getenv("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "DCO_Expedientes" / "feedback_sesion.json"


class FeedbackAPI:
    """Crear una instancia para la app de usuarios y otra para DCO_Admin.

    Las peticiones de red deben llamarse desde un hilo, nunca desde el hilo
    gráfico de PySide6. No almacena contraseña de administrador.
    """
    _lock_archivo = threading.RLock()

    def __init__(self, admin=False):
        if not configurado():
            raise FeedbackError("DCO Feedback todavía no está configurado (feedback_config.py).")
        self.admin = admin
        self.session = requests.Session()
        self.sesion = {}
        if not admin:
            with self._lock_archivo:
                try:
                    self.sesion = json.loads(_ruta_sesion().read_text(encoding="utf-8"))
                except (FileNotFoundError, ValueError, OSError):
                    pass

    def _headers(self, *, autenticado=True):
        # Una clave sb_publishable_ NO es un JWT: solo se manda en apikey.
        # Authorization se envía exclusivamente cuando existe un token de sesión.
        encabezados = {
            "apikey": SUPABASE_PUBLIC_KEY,
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        token = self.sesion.get("access_token") if autenticado else None
        if token:
            encabezados["Authorization"] = f"Bearer {token}"
        return encabezados

    def _request(self, metodo, ruta, *, autenticado=True, **kwargs):
        try:
            encabezados = self._headers(autenticado=autenticado)
            encabezados.update(kwargs.pop("headers", {}))
            r = self.session.request(
                metodo, SUPABASE_URL + ruta,
                headers=encabezados,
                timeout=(8, 16), **kwargs
            )
            r.raise_for_status()
            if not r.content:
                return None
            return r.json()
        except requests.exceptions.HTTPError as exc:
            try:
                detalle = exc.response.json()
                mensaje = detalle.get("message") or detalle.get("msg") or detalle.get("error_description") or detalle.get("error")
            except Exception:
                mensaje = None
            codigo = exc.response.status_code if exc.response is not None else "?"
            raise FeedbackError(f"Error del servidor ({codigo}): {mensaje or 'revisa la configuración o permisos.'}") from exc
        except (requests.exceptions.RequestException, ValueError) as exc:
            raise FeedbackError("No se pudo conectar al servicio de encuestas.") from exc

    def _guardar_sesion(self):
        if self.admin:
            return
        with self._lock_archivo:
            ruta = _ruta_sesion()
            ruta.parent.mkdir(parents=True, exist_ok=True)
            temporal = ruta.with_suffix(".tmp")
            contenido = {k: self.sesion.get(k) for k in ("access_token", "refresh_token", "expires_at", "user_id")}
            with temporal.open("w", encoding="utf-8") as f:
                json.dump(contenido, f)
            if os.name != "nt":
                os.chmod(temporal, 0o600)
            os.replace(temporal, ruta)

    def _actualizar_sesion(self, datos):
        # Los endpoints Auth devuelven tokens en nivel raíz.
        self.sesion = {
            "access_token": datos["access_token"],
            "refresh_token": datos["refresh_token"],
            "expires_at": time.time() + int(datos.get("expires_in", 3600)),
            "user_id": (datos.get("user") or {}).get("id") or self.sesion.get("user_id"),
        }
        self._guardar_sesion()

    def asegurar_login(self):
        if self.sesion.get("access_token") and self.sesion.get("expires_at", 0) > time.time() + 60:
            return
        if self.sesion.get("refresh_token"):
            try:
                datos = self._request(
                    "POST", "/auth/v1/token?grant_type=refresh_token", autenticado=False,
                    json={"refresh_token": self.sesion["refresh_token"]},
                )
                self._actualizar_sesion(datos)
                return
            except FeedbackError:
                if self.admin:
                    raise
        if self.admin:
            raise FeedbackError("La sesión administrativa caducó. Vuelve a iniciar sesión.")
        # POST signup sin email crea usuario anónimo si la opción está habilitada.
        datos = self._request(
            "POST", "/auth/v1/signup", autenticado=False,
            json={"data": {}, "gotrue_meta_security": {"captcha_token": None}},
        )
        if not datos or not datos.get("access_token"):
            raise FeedbackError("Activa 'Allow anonymous sign-ins' en Supabase Auth.")
        self._actualizar_sesion(datos)

    def login_admin(self, email, password):
        if not self.admin:
            raise FeedbackError("Esta operación está reservada al administrador.")
        datos = self._request(
            "POST", "/auth/v1/token?grant_type=password", autenticado=False,
            json={"email": email.strip(), "password": password},
        )
        self._actualizar_sesion(datos)
        filas = self._request(
            "GET", "/rest/v1/feedback_admins?select=user_id&limit=1"
        )
        if not filas:
            self.sesion = {}
            raise FeedbackError("La cuenta autenticó, pero no tiene permiso de administrador.")
        return True

    def campaña_activa(self):
        self.asegurar_login()
        filas = self._request(
            "GET", "/rest/v1/feedback_campaigns?select=id,title,description,created_at"
            "&active=eq.true&order=created_at.desc&limit=1"
        ) or []
        if not filas:
            return None
        campaña = filas[0]
        # Solo un voto por campaña y cuenta anónima del equipo.
        ya = self._request(
            "GET", "/rest/v1/feedback_responses?select=id"
            "&campaign_id=eq." + quote(str(campaña["id"])) + "&limit=1"
        ) or []
        return None if ya else campaña

    def enviar_respuesta(self, campaign_id, rating, liked, improve, category):
        if not (1 <= int(rating) <= 5):
            raise FeedbackError("Selecciona una calificación de 1 a 5.")
        self.asegurar_login()
        cuerpo = {
            "campaign_id": str(campaign_id),
            "user_id": self.sesion["user_id"],
            "rating": int(rating),
            "liked": liked.strip()[:2000],
            "improve": improve.strip()[:2000],
            "category": category.strip()[:80],
        }
        return self._request("POST", "/rest/v1/feedback_responses", json=cuerpo)

    def enviar_sugerencia(self, texto):
        texto = texto.strip()
        if len(texto) < 5:
            raise FeedbackError("La sugerencia debe tener al menos cinco caracteres.")
        self.asegurar_login()
        return self._request(
            "POST", "/rest/v1/feedback_suggestions",
            json={"user_id": self.sesion["user_id"], "body": texto[:2000]},
        )

    def listar_campañas(self):
        self.asegurar_login()
        return self._request(
            "GET", "/rest/v1/feedback_campaigns?select=id,title,description,active,created_at"
            "&order=created_at.desc&limit=100"
        ) or []

    def crear_campaña(self, titulo, descripcion):
        self.asegurar_login()
        if not titulo.strip():
            raise FeedbackError("Escribe el título de la campaña.")
        filas = self._request(
            "POST", "/rest/v1/feedback_campaigns",
            json={"title": titulo.strip()[:120], "description": descripcion.strip()[:4000],
                  "active": False, "created_by": self.sesion["user_id"]},
            params={"select": "id,title,description,active,created_at"},
            headers={**self._headers(), "Prefer": "return=representation"},
        )
        return filas[0] if filas else None

    def activar_campaña(self, campaign_id):
        self.asegurar_login()
        # La RPC desactiva otras campañas en una transacción.
        return self._request(
            "POST", "/rest/v1/rpc/feedback_activate_campaign",
            json={"p_campaign_id": str(campaign_id)},
        )

    def detener_campaña(self):
        self.asegurar_login()
        return self._request("POST", "/rest/v1/rpc/feedback_stop_campaign", json={})

    def listar_respuestas(self, campaign_id):
        self.asegurar_login()
        return self._request(
            "GET", "/rest/v1/feedback_responses?select=rating,liked,improve,category,created_at"
            "&campaign_id=eq." + quote(str(campaign_id)) + "&order=created_at.desc&limit=1000"
        ) or []

    def listar_sugerencias(self):
        self.asegurar_login()
        return self._request(
            "GET", "/rest/v1/feedback_suggestions?select=body,created_at"
            "&order=created_at.desc&limit=1000"
        ) or []
