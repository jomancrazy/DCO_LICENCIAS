"""Actualizador de DCO Expedientes para versiones empaquetadas Windows/macOS.
Requiere repositorio y releases públicos de GitHub.
"""
import os
import re
import sys
import time
import json
import shutil
import zipfile
import tempfile
import subprocess
from pathlib import Path
from urllib.parse import quote

import requests
from PySide6.QtCore import QThread, Signal, QEventLoop, Qt
from PySide6.QtWidgets import QMessageBox, QProgressDialog

# Cambiar en CADA nueva compilación publicada.
VERSION_ACTUAL = "1.0.3"
REPOSITORIO = "jomancrazy/DCO_LICENCIAS"
VERSION_URL = f"https://raw.githubusercontent.com/{REPOSITORIO}/main/version.txt"
NOMBRE_ZIP_WINDOWS = "DCO_Expedientes_Windows.zip"
NOMBRE_ZIP_MAC = "DCO_Expedientes_Mac.zip"
NOMBRE_EXE_WINDOWS = "DCO_Expedientes.exe"
NOMBRE_APP_MAC = "DCO_Expedientes.app"
LOG = Path.home() / "dco_actualizador.log"


def registrar(mensaje):
    try:
        with LOG.open("a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {mensaje}\n")
    except OSError:
        pass


def normalizar_version(valor):
    valor = valor.strip().lstrip("\ufeff").removeprefix("v").removeprefix("V")
    if not re.fullmatch(r"\d+(?:\.\d+){1,3}", valor):
        raise ValueError(f"Versión inválida: {valor!r}")
    return tuple(map(int, valor.split(".")))


def plataforma():
    if sys.platform == "darwin":
        return "mac"
    if sys.platform == "win32":
        return "windows"
    raise RuntimeError("Solo compatible con Windows y macOS")


def ruta_instalada():
    if not getattr(sys, "frozen", False):
        raise RuntimeError("La instalación automática solo funciona en la app compilada (.exe/.app).")
    executable = Path(sys.executable).absolute()
    if plataforma() == "windows":
        if not executable.is_file():
            raise RuntimeError("No se encontró el ejecutable instalado")
        return executable
    for p in executable.parents:
        if p.suffix.lower() == ".app":
            return p
    raise RuntimeError("No se encontró el paquete .app instalado")


def consultar_version_remota():
    response = requests.get(VERSION_URL, timeout=15, headers={"Cache-Control": "no-cache"},
                            params={"t": time.time_ns()})
    response.raise_for_status()
    version = response.text.strip().lstrip("\ufeff")
    normalizar_version(version)
    return version


def obtener_asset_de_version(version, nombre):
    """Busca solo el archivo del Release etiquetado con la versión solicitada."""
    v = ".".join(str(n) for n in normalizar_version(version))
    sesion = requests.Session()
    sesion.headers.update({"Accept": "application/vnd.github+json", "User-Agent": "DCO-Updater"})
    for tag in (f"v{v}", v):
        respuesta = sesion.get(
            f"https://api.github.com/repos/{REPOSITORIO}/releases/tags/{quote(tag)}",
            timeout=20)
        if respuesta.status_code == 404:
            continue
        respuesta.raise_for_status()
        release = respuesta.json()
        if release.get("draft") or release.get("prerelease"):
            raise RuntimeError(f"El Release {tag} todavía no está publicado como versión estable")
        for asset in release.get("assets", []):
            if asset.get("name") == nombre:
                url = asset.get("browser_download_url", "")
                if not url.startswith("https://github.com/"):
                    raise RuntimeError("URL de descarga inesperada")
                return url
        raise RuntimeError(f"El Release {tag} no tiene el archivo {nombre}")
    raise RuntimeError(f"No existe un Release de GitHub con etiqueta v{v} o {v}")


class DescargaThread(QThread):
    progreso = Signal(int)
    terminado = Signal(bool, str)

    def __init__(self, url, destino):
        super().__init__()
        self.url = url
        self.destino = Path(destino)

    def run(self):
        parcial = self.destino.with_suffix(".part")
        try:
            with requests.get(self.url, stream=True, timeout=(15, 90),
                              headers={"User-Agent": "DCO-Updater"}) as r:
                r.raise_for_status()
                total = int(r.headers.get("Content-Length", "0"))
                recibido = 0
                with parcial.open("wb") as f:
                    for bloque in r.iter_content(1024 * 128):
                        if self.isInterruptionRequested():
                            raise InterruptedError("Descarga cancelada")
                        if bloque:
                            f.write(bloque)
                            recibido += len(bloque)
                            if total:
                                self.progreso.emit(min(99, int(recibido * 100 / total)))
                if total and recibido != total:
                    raise RuntimeError("La descarga quedó incompleta")
            if not zipfile.is_zipfile(parcial):
                raise RuntimeError("El archivo descargado no es un ZIP válido")
            with zipfile.ZipFile(parcial) as z:
                corrupto = z.testzip()
                if corrupto:
                    raise RuntimeError(f"Archivo corrupto en ZIP: {corrupto}")
            parcial.replace(self.destino)
            self.progreso.emit(100)
            self.terminado.emit(True, "")
        except Exception as e:
            parcial.unlink(missing_ok=True)
            registrar(f"Error al descargar: {e}")
            self.terminado.emit(False, str(e))


def descargar(url, archivo, parent):
    dialogo = QProgressDialog("Descargando actualización...", "Cancelar", 0, 100, parent)
    dialogo.setWindowTitle("DCO Expedientes - Actualización")
    dialogo.setMinimumDuration(0)
    dialogo.setWindowModality(Qt.WindowModality.ApplicationModal)
    dialogo.setAutoClose(False)
    hilo = DescargaThread(url, archivo)
    loop = QEventLoop()
    resultado = {"ok": False, "error": ""}

    def terminado(ok, error):
        resultado.update(ok=ok, error=error)
        dialogo.close()
        loop.quit()

    hilo.progreso.connect(dialogo.setValue)
    hilo.terminado.connect(terminado)
    dialogo.canceled.connect(hilo.requestInterruption)
    hilo.start()
    loop.exec()
    hilo.wait()
    return resultado["ok"], resultado["error"]


def validar_zip(zip_path, plat):
    """Requiere el ejecutable o bundle en la raíz, sin carpetas envolventes."""
    with zipfile.ZipFile(zip_path) as z:
        nombres = z.namelist()
        for nombre in nombres:
            p = Path(nombre)
            if p.is_absolute() or ".." in p.parts or "\\" in nombre:
                raise RuntimeError(f"Ruta peligrosa dentro del ZIP: {nombre}")
        if plat == "windows":
            if NOMBRE_EXE_WINDOWS not in nombres:
                raise RuntimeError(f"El ZIP Windows debe contener {NOMBRE_EXE_WINDOWS} en la raíz")
        elif not any(x.startswith(NOMBRE_APP_MAC + "/Contents/MacOS/") for x in nombres):
            raise RuntimeError(f"El ZIP Mac debe contener {NOMBRE_APP_MAC} en la raíz")


def instalar_windows(zip_path, ejecutable):
    destino = ejecutable.parent
    # Script externo espera que termine la aplicación antes de reemplazar el EXE.
    bat = Path(tempfile.mkdtemp(prefix="dco_updater_")) / "actualizar.bat"
    contenido = r'''@echo off
setlocal
set "PID=%~1"
set "ZIP=%~2"
set "DEST=%~3"
set "EXE=%~4"
set "STAGE=%TEMP%\dco_stage_%RANDOM%%RANDOM%"
set /a ATTEMPTS=0
:WAIT
 tasklist /FI "PID eq %PID%" 2>NUL | findstr /R /C:"[ ]%PID%[ ]" >NUL
 if not errorlevel 1 (
   timeout /t 1 /nobreak >NUL
   set /a ATTEMPTS+=1
   if %ATTEMPTS% GEQ 120 goto ERROR
   goto WAIT
 )
 powershell -NoProfile -Command "Expand-Archive -LiteralPath $env:ZIP -DestinationPath $env:STAGE -Force"
 if errorlevel 1 goto ERROR
 if not exist "%STAGE%\%EXE%" goto ERROR
 robocopy "%STAGE%" "%DEST%" /E /R:3 /W:2 >NUL
 if %ERRORLEVEL% GEQ 8 goto ERROR
 start "" "%DEST%\%EXE%"
 del "%ZIP%" >NUL 2>&1
 rmdir /S /Q "%STAGE%" >NUL 2>&1
 exit /b 0
:ERROR
 echo No se pudo instalar la actualizacion. Revisa permisos y archivos.
 pause
 exit /b 1
'''
    bat.write_text(contenido, encoding="utf-8")
    subprocess.Popen(["cmd.exe", "/c", str(bat), str(os.getpid()),
                      str(zip_path), str(destino), ejecutable.name],
                     creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0), close_fds=True)


def instalar_mac(zip_path, app_actual):
    # La aplicación existente puede tener un nombre diferente al del ZIP.
    temporal = Path(tempfile.mkdtemp(prefix="dco_updater_"))
    script = temporal / "instalar.sh"
    registro = Path.home() / "dco_actualizador_instalacion.log"
    script.write_text(r'''#!/bin/bash
set -u
PID="$1"
ZIP="$2"
ACTUAL="$3"
STAGE="$4"
LOG="$5"
exec >> "$LOG" 2>&1
for i in {1..120}; do
  if ! kill -0 "$PID" 2>/dev/null; then break; fi
  sleep 1
done
if kill -0 "$PID" 2>/dev/null; then echo "La aplicación no terminó"; exit 1; fi
mkdir -p "$STAGE/extracted" || exit 1
/usr/bin/ditto -x -k "$ZIP" "$STAGE/extracted" || exit 1
NUEVA="$STAGE/extracted/DCO_Expedientes.app"
BACKUP="${ACTUAL}.backup"
if [ ! -d "$NUEVA" ]; then echo "No se encontró el nuevo .app"; exit 1; fi
# No continuar si no hay permisos de escritura en la carpeta instalada.
if [ ! -w "$(dirname "$ACTUAL")" ]; then echo "No hay permisos en carpeta de instalación"; exit 1; fi
rm -rf "$BACKUP" || exit 1
mv "$ACTUAL" "$BACKUP" || exit 1
if ! mv "$NUEVA" "$ACTUAL"; then
  mv "$BACKUP" "$ACTUAL"
  echo "Falló la instalación; se recuperó la versión anterior"
  exit 1
fi
open "$ACTUAL"
rm -rf "$BACKUP"
rm -f "$ZIP"
rm -rf "$STAGE"
''', encoding="utf-8")
    script.chmod(0o700)
    subprocess.Popen(["/bin/bash", str(script), str(os.getpid()), str(zip_path),
                      str(app_actual), str(temporal), str(registro)],
                     start_new_session=True, close_fds=True)


def verificar_e_instalar_actualizacion(parent=None):
    """Devuelve True solo si se lanzó un instalador y hay que salir del programa."""
    try:
        registrar(f"Inicio: local={VERSION_ACTUAL} plataforma={sys.platform} ejecutable={sys.executable}")
        try:
            remota = consultar_version_remota()
        except Exception as e:
            registrar(f"No se pudo consultar version.txt: {e}")
            return False  # no bloquear el trabajo si GitHub falla
        registrar(f"Remota={remota}")
        if normalizar_version(remota) <= normalizar_version(VERSION_ACTUAL):
            registrar("Sin actualizaciones")
            return False
        plat = plataforma()
        nombre = NOMBRE_ZIP_WINDOWS if plat == "windows" else NOMBRE_ZIP_MAC
        url = obtener_asset_de_version(remota, nombre)
        registrar(f"Asset: {url}")
        decision = QMessageBox.question(parent, "Nueva versión de DCO Expedientes",
            f"Versión instalada: {VERSION_ACTUAL}\nDisponible: {remota}\n\n"
            "¿Descargar e instalar ahora? Se cerrará el programa para completar la actualización.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes)
        if decision != QMessageBox.StandardButton.Yes:
            return False
        ejecutable = ruta_instalada()
        ruta_zip = Path(tempfile.gettempdir()) / f"dco_update_{os.getpid()}.zip"
        ok, error = descargar(url, ruta_zip, parent)
        if not ok:
            QMessageBox.warning(parent, "Actualización cancelada o fallida", error)
            return False
        validar_zip(ruta_zip, plat)
        if plat == "windows":
            instalar_windows(ruta_zip, ejecutable)
        else:
            instalar_mac(ruta_zip, ejecutable)
        registrar("Instalador externo iniciado")
        QMessageBox.information(parent, "Actualización preparada",
            "La descarga terminó. La aplicación se cerrará para instalar la actualización.")
        return True
    except Exception as e:
        registrar(f"Error general: {type(e).__name__}: {e}")
        QMessageBox.critical(parent, "Error de actualización",
            f"No se pudo actualizar DCO Expedientes:\n\n{e}\n\nRegistro: {LOG}")
        return False