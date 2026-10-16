
import os
import sys
import re
import time
import shutil
import zipfile
import tempfile
import subprocess
import requests

from pathlib import Path
from PySide6.QtCore import QThread, Signal, QEventLoop
from PySide6.QtWidgets import (
    QApplication,
    QMessageBox,
    QProgressDialog,
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

VERSION_ACTUAL = "1.0.0"

VERSION_URL = (
    "https://raw.githubusercontent.com/"
    "jomancrazy/DCO_LICENCIAS/main/version.txt"
)

URL_WINDOWS = (
    "https://github.com/jomancrazy/DCO_LICENCIAS/"
    "releases/latest/download/DCO_Expedientes_Windows.zip"
)

URL_MAC = (
    "https://github.com/jomancrazy/DCO_LICENCIAS/"
    "releases/latest/download/DCO_Expedientes_Mac.zip"
)

NOMBRE_APP_MAC = "DCO_Expedientes.app"
NOMBRE_EXE_WINDOWS = "DCO_Expedientes.exe"


# ============================================================
# REGISTRO DE ERRORES
# ============================================================

def registrar(mensaje):
    """Guarda los mensajes de diagnóstico en un archivo local."""
    try:
        ruta = Path.home() / "dco_actualizador.log"

        with ruta.open("a", encoding="utf-8") as archivo:
            fecha = time.strftime("%Y-%m-%d %H:%M:%S")
            archivo.write(f"[{fecha}] {mensaje}\n")

    except Exception:
        pass


# ============================================================
# DETECCIÓN DE PLATAFORMA Y RUTAS
# ============================================================

def obtener_plataforma():
    if sys.platform == "darwin":
        return "mac"

    if sys.platform.startswith("win"):
        return "windows"

    raise RuntimeError(
        f"Sistema operativo no compatible: {sys.platform}"
    )


def obtener_app_mac():
    """Localiza el bundle .app que contiene al ejecutable."""
    ejecutable = Path(sys.executable).resolve()

    for carpeta in ejecutable.parents:
        if carpeta.name.endswith(".app"):
            return carpeta

    # Alternativa para cuando se ejecuta desde Python.
    candidato = Path(sys.argv[0]).resolve()

    for carpeta in (candidato, *candidato.parents):
        if carpeta.name.endswith(".app"):
            return carpeta

    raise RuntimeError(
        "No se pudo localizar DCO_Expedientes.app."
    )


def obtener_directorio_instalacion():
    plataforma = obtener_plataforma()

    if plataforma == "mac":
        app = obtener_app_mac()
        return app.parent

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(sys.argv[0]).resolve().parent


# ============================================================
# COMPARACIÓN DE VERSIONES
# ============================================================

def normalizar_version(version):
    """
    Acepta versiones como 1.0.1, 1.1.0 o v1.1.0.
    Rechaza contenido inesperado en version.txt.
    """
    version = version.strip().lstrip("\ufeff")

    if version.lower().startswith("v"):
        version = version[1:]

    if not re.fullmatch(r"\d+(?:\.\d+)*", version):
        raise ValueError(
            f"Formato de versión no válido: {version!r}"
        )

    return tuple(int(x) for x in version.split("."))


def hay_version_nueva(actual, remota):
    actual = normalizar_version(actual)
    remota = normalizar_version(remota)

    longitud = max(len(actual), len(remota))

    actual += (0,) * (longitud - len(actual))
    remota += (0,) * (longitud - len(remota))

    return remota > actual


# ============================================================
# DESCARGA EN SEGUNDO PLANO
# ============================================================

class DescargaThread(QThread):
    progreso = Signal(int)
    terminado = Signal(bool, str)

    def __init__(self, url, destino):
        super().__init__()
        self.url = url
        self.destino = destino

    def run(self):
        temporal = self.destino + ".part"

        try:
            with requests.get(
                self.url,
                stream=True,
                timeout=(15, 60),
                headers={
                    "Cache-Control": "no-cache",
                    "User-Agent": "DCO-Expedientes-Updater",
                },
                allow_redirects=True,
            ) as respuesta:

                respuesta.raise_for_status()

                total = int(
                    respuesta.headers.get("content-length", 0)
                )
                descargado = 0

                with open(temporal, "wb") as archivo:
                    for bloque in respuesta.iter_content(
                        chunk_size=64 * 1024
                    ):
                        if self.isInterruptionRequested():
                            raise InterruptedError(
                                "Descarga cancelada."
                            )

                        if not bloque:
                            continue

                        archivo.write(bloque)
                        descargado += len(bloque)

                        if total > 0:
                            porcentaje = min(
                                100,
                                int(descargado * 100 / total),
                            )
                            self.progreso.emit(porcentaje)

            # No aceptar una descarga vacía o incompleta.
            if not os.path.isfile(temporal):
                raise RuntimeError("No se generó el archivo ZIP.")

            if os.path.getsize(temporal) == 0:
                raise RuntimeError("El ZIP descargado está vacío.")

            if not zipfile.is_zipfile(temporal):
                raise RuntimeError(
                    "La descarga no es un ZIP válido. "
                    "Comprueba los archivos del GitHub Release."
                )

            os.replace(temporal, self.destino)
            self.terminado.emit(True, "")

        except Exception as error:
            try:
                if os.path.exists(temporal):
                    os.remove(temporal)
            except OSError:
                pass

            registrar(f"Error de descarga: {error}")
            self.terminado.emit(False, str(error))


# ============================================================
# CONSULTAR LA VERSIÓN REMOTA
# ============================================================

def consultar_version_remota():
    respuesta = requests.get(
        VERSION_URL,
        params={"_": str(time.time_ns())},
        timeout=10,
        headers={
            "Cache-Control": "no-cache, no-store",
            "Pragma": "no-cache",
            "User-Agent": "DCO-Expedientes-Updater",
        },
    )

    respuesta.raise_for_status()

    version = respuesta.text.strip().lstrip("\ufeff")

    # Validar antes de usar el dato.
    normalizar_version(version)

    return version


# ============================================================
# INSTALADOR AUXILIAR PARA WINDOWS
# ============================================================

def crear_instalador_windows(zip_path, destino):
    exe = destino / NOMBRE_EXE_WINDOWS

    if not exe.is_file():
        raise RuntimeError(
            f"No se encontró el ejecutable instalado: {exe}"
        )

    temporal = Path(tempfile.gettempdir())
    script = temporal / f"dco_update_{os.getpid()}.bat"

    # El ZIP debe contener DCO_Expedientes.exe en la raíz.
    # El instalador espera que el proceso actual termine.
    contenido = r"""@echo off
setlocal
set "PID=%~1"
set "ZIP=%~2"
set "DEST=%~3"
set "EXE=%~4"
set "STAGE=%TEMP%\DCO_Update_%RANDOM%%RANDOM%"

:WAIT_APP
tasklist /FI "PID eq %PID%" 2>NUL | findstr /R /C:"[ ]%PID%[ ]" >NUL
if not errorlevel 1 (
    timeout /t 1 /nobreak >NUL
    goto WAIT_APP
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "Expand-Archive -LiteralPath $env:ZIP -DestinationPath $env:STAGE -Force"
if errorlevel 1 goto ERROR

if not exist "%STAGE%\%EXE%" goto ERROR

robocopy "%STAGE%" "%DEST%" /E /R:3 /W:1
if %ERRORLEVEL% GEQ 8 goto ERROR

start "" "%DEST%\%EXE%"
del "%ZIP%" >NUL 2>&1
rmdir /S /Q "%STAGE%" >NUL 2>&1
del "%~f0"
exit /b 0

:ERROR
echo Error instalando DCO Expedientes. Revisa el ZIP y los permisos.
pause
exit /b 1
"""

    # Pasar rutas por variables de entorno evita problemas
    # con espacios y caracteres especiales en los argumentos.
    contenido = contenido.replace(
        '"%env:ZIP%"', '"%ZIP%"'
    ).replace(
        '"%env:STAGE%"', '"%STAGE%"'
    )

    script.write_text(contenido, encoding="utf-8")

    # BAT recibe PID, ZIP, destino y nombre del ejecutable.
    subprocess.Popen(
        [
            "cmd.exe",
            "/c",
            str(script),
            str(os.getpid()),
            str(zip_path),
            str(destino),
            NOMBRE_EXE_WINDOWS,
        ],
        creationflags=getattr(
            subprocess, "CREATE_NEW_CONSOLE", 0
        ),
        close_fds=True,
    )


# ============================================================
# INSTALADOR AUXILIAR PARA MACOS
# ============================================================

def crear_instalador_mac(zip_path, app_actual):
    app_actual = Path(app_actual).resolve()
    destino = app_actual.parent
    app_nueva = destino / NOMBRE_APP_MAC

    if not app_actual.is_dir():
        raise RuntimeError(
            f"No se encontró la aplicación instalada: {app_actual}"
        )

    temporal = Path(tempfile.mkdtemp(prefix="dco_update_"))
    script = temporal / "instalar.command"
    log = Path.home() / "dco_actualizador_instalacion.log"

    # Las rutas se pasan como argumentos, no se interpolan
    # directamente en el código del shell.
    contenido = r"""#!/bin/bash
set -u

PID="$1"
ZIP="$2"
DEST="$3"
APP_NAME="$4"
STAGE="$5"
LOG="$6"

exec >> "$LOG" 2>&1

echo "Esperando el cierre de DCO Expedientes..."
while kill -0 "$PID" 2>/dev/null; do
    sleep 1
done

mkdir -p "$STAGE/extracted"

if ! /usr/bin/ditto -x -k "$ZIP" "$STAGE/extracted"; then
    echo "Error extrayendo ZIP"
    exit 1
fi

NEW_APP="$STAGE/extracted/$APP_NAME"
OLD_APP="$DEST/$APP_NAME"
BACKUP="$DEST/$APP_NAME.backup"

if [ ! -d "$NEW_APP" ]; then
    echo "No se encontró $APP_NAME en la raíz del ZIP"
    exit 1
fi

if [ -e "$BACKUP" ]; then
    rm -rf "$BACKUP"
fi

if [ -e "$OLD_APP" ]; then
    mv "$OLD_APP" "$BACKUP" || exit 1
fi

if ! mv "$NEW_APP" "$OLD_APP"; then
    echo "Error instalando. Restaurando copia anterior."
    if [ -e "$BACKUP" ]; then
        mv "$BACKUP" "$OLD_APP"
    fi
    exit 1
fi

rm -rf "$BACKUP"
open "$OLD_APP"
rm -f "$ZIP"
rm -rf "$STAGE"
"""

    script.write_text(contenido, encoding="utf-8")
    script.chmod(0o700)

    subprocess.Popen(
        [
            "/bin/bash",
            str(script),
            str(os.getpid()),
            str(zip_path),
            str(destino),
            NOMBRE_APP_MAC,
            str(temporal),
            str(log),
        ],
        start_new_session=True,
        close_fds=True,
    )


# ============================================================
# DESCARGAR Y APLICAR ACTUALIZACIÓN
# ============================================================

def descargar_actualizacion(url, destino, parent=None):
    progreso = QProgressDialog(
        "Descargando actualización de DCO Expedientes...",
        "Cancelar",
        0,
        100,
        parent,
    )

    progreso.setWindowTitle("Actualización de DCO Expedientes")
    progreso.setMinimumDuration(0)
    progreso.setAutoClose(False)
    progreso.setAutoReset(False)
    progreso.setWindowModality(
        progreso.windowModality()
        if parent is None
        else 2
    )

    thread = DescargaThread(url, str(destino))
    loop = QEventLoop()

    resultado = {"ok": False, "error": ""}

    def actualizar_porcentaje(valor):
        progreso.setValue(valor)

    def finalizar(ok, error):
        resultado["ok"] = ok
        resultado["error"] = error
        progreso.close()
        loop.quit()

    def cancelar():
        if thread.isRunning():
            thread.requestInterruption()
        progreso.setLabelText("Cancelando descarga...")

    thread.progreso.connect(actualizar_porcentaje)
    thread.terminado.connect(finalizar)
    progreso.canceled.connect(cancelar)

    thread.start()
    loop.exec()

    # No dejar el hilo escribiendo después de salir.
    thread.wait()

    return resultado["ok"], resultado["error"]


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def verificar_e_instalar_actualizacion(parent=None):
    """
    Comprueba la versión remota y, si es más nueva,
    descarga e inicia el instalador correspondiente.
    """
    try:
        registrar(
            f"Inicio. Actual={VERSION_ACTUAL}; "
            f"plataforma={sys.platform}; "
            f"ejecutable={sys.executable}"
        )

        try:
            version_remota = consultar_version_remota()
        except Exception as error:
            registrar(f"Error consultando GitHub: {error}")
            # Si GitHub no responde, permitir abrir la aplicación.
            return False

        registrar(f"Versión remota={version_remota}")

        if not hay_version_nueva(
            VERSION_ACTUAL,
            version_remota,
        ):
            registrar("No hay una versión más nueva.")
            return False

        respuesta = QMessageBox.question(
            parent,
            "Actualización disponible",
            (
                f"Tu versión: {VERSION_ACTUAL}\n"
                f"Nueva versión: {version_remota}\n\n"
                "¿Deseas descargar e instalar la actualización?\n"
                "La aplicación se cerrará para aplicar los cambios."
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )

        if respuesta != QMessageBox.StandardButton.Yes:
            return False

        plataforma = obtener_plataforma()
        url = URL_WINDOWS if plataforma == "windows" else URL_MAC

        temporal = Path(tempfile.gettempdir())
        zip_path = temporal / (
            f"dco_update_{os.getpid()}.zip"
        )

        ok, error = descargar_actualizacion(
            url,
            zip_path,
            parent,
        )

        if not ok:
            QMessageBox.critical(
                parent,
                "Error de actualización",
                f"No se pudo descargar la actualización:\n{error}",
            )
            return False

        # Validación adicional de estructura antes de instalar.
        with zipfile.ZipFile(zip_path, "r") as archivo_zip:
            nombres = archivo_zip.namelist()

            if plataforma == "mac":
                valido = any(
                    nombre.startswith(NOMBRE_APP_MAC + "/")
                    for nombre in nombres
                )
            else:
                valido = NOMBRE_EXE_WINDOWS in nombres

            if not valido:
                raise RuntimeError(
                    "El ZIP no contiene la aplicación esperada "
                    "en su raíz. Revisa cómo se creó el Release."
                )

        if plataforma == "mac":
            app_actual = obtener_app_mac()
            crear_instalador_mac(zip_path, app_actual)
        else:
            destino = obtener_directorio_instalacion()
            crear_instalador_windows(zip_path, destino)

        registrar("Instalador auxiliar iniciado.")

        QMessageBox.information(
            parent,
            "Actualización preparada",
            (
                "La actualización se ha descargado y validado.\n\n"
                "DCO Expedientes se cerrará para aplicar los cambios "
                "y volverá a abrirse si la instalación termina bien."
            ),
        )

        # El proceso auxiliar esperará a que la aplicación termine.
        QApplication.quit()
        return True

    except Exception as error:
        registrar(f"Error general del actualizador: {error}")

        QMessageBox.critical(
            parent,
            "Error de actualización",
            f"No se pudo preparar la actualización:\n\n{error}",
        )

        return False
