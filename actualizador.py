import hashlib
import os
import re
import sys
import time
import zipfile
import tempfile
import subprocess
from pathlib import Path, PurePosixPath
from urllib.parse import quote

import requests
from PySide6.QtCore import QObject, QThread, Signal, Qt, QTimer
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QApplication,
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

# La versión de CADA compilación debe ser la que tenía al publicarse.
# Ejemplo: para crear el ejecutable 1.0.3, cambiar a "1.0.3"
# ANTES de compilarlo; la compilación 1.0.2 debe seguir diciendo 1.0.2.
VERSION_ACTUAL = "1.0.6"

# Comprobar actualizaciones mientras la aplicación está abierta.
INTERVALO_REVISION_MS = 30 * 60 * 1000  # 30 minutos

REPOSITORIO = "jomancrazy/DCO_LICENCIAS"
VERSION_URL = f"https://raw.githubusercontent.com/{REPOSITORIO}/main/version.txt"
API_RELEASES = f"https://api.github.com/repos/{REPOSITORIO}/releases/tags"

ZIP_WINDOWS = "DCO_Expedientes_Windows.zip"
ZIP_MAC = "DCO_Expedientes_Mac.zip"
EXE_WINDOWS = "DCO_Expedientes.exe"
APP_MAC = "DCO_Expedientes.app"

LOG = Path.home() / "dco_actualizador.log"
HEADERS_GITHUB = {
    "Accept": "application/vnd.github+json",
    "User-Agent": "DCO-Expedientes-Actualizador",
}


def registrar(mensaje):
    try:
        with LOG.open("a", encoding="utf-8") as archivo:
            ahora = time.strftime("%Y-%m-%d %H:%M:%S")
            archivo.write(f"[{ahora}] {mensaje}\n")
    except OSError:
        pass


def normalizar_version(version):
    texto = str(version).strip().lstrip("\ufeff")
    if texto.lower().startswith("v"):
        texto = texto[1:]
    if not re.fullmatch(r"\d+(?:\.\d+){1,3}", texto):
        raise ValueError(f"Versión no válida: {version!r}")
    return tuple(int(parte) for parte in texto.split("."))


def es_version_nueva(local, remota):
    a, b = normalizar_version(local), normalizar_version(remota)
    longitud = max(len(a), len(b))
    return b + (0,) * (longitud - len(b)) > a + (0,) * (longitud - len(a))


def plataforma_actual():
    if sys.platform == "darwin":
        return "mac"
    if sys.platform == "win32":
        return "windows"
    raise RuntimeError(f"Sistema operativo no compatible: {sys.platform}")


def obtener_instalacion():
    """Localiza el ejecutable instalado. No sirve en modo Python de desarrollo."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "La instalación automática solo funciona en la aplicación compilada "
            "(.exe en Windows o .app en macOS)."
        )

    ruta = Path(sys.executable).resolve()
    if plataforma_actual() == "windows":
        if not ruta.is_file():
            raise RuntimeError(f"No se encontró el ejecutable: {ruta}")
        return ruta

    for carpeta in ruta.parents:
        if carpeta.name.lower().endswith(".app"):
            if not os.access(carpeta.parent, os.W_OK):
                raise RuntimeError(
                    "No hay permisos para reemplazar la aplicación Mac en "
                    f"{carpeta.parent}. Copia primero la app a una carpeta "
                    "donde tengas permisos de escritura y ábrela desde ahí."
                )
            return carpeta
    raise RuntimeError("No se encontró el paquete .app instalado.")


# ============================================================
# CONSULTA DE GITHUB
# ============================================================

def consultar_version_remota():
    respuesta = requests.get(
        VERSION_URL,
        params={"nocache": str(time.time_ns())},
        headers={**HEADERS_GITHUB, "Cache-Control": "no-cache, no-store"},
        timeout=(10, 20),
    )
    respuesta.raise_for_status()
    version = respuesta.text.strip().lstrip("\ufeff")
    normalizar_version(version)
    return version


def obtener_release_y_archivo(version, nombre_zip):
    """Obtiene (url_zip, notas_markdown, sha256_o_None) del Release exacto."""
    version_normalizada = ".".join(map(str, normalizar_version(version)))

    for etiqueta in (f"v{version_normalizada}", version_normalizada):
        respuesta = requests.get(
            f"{API_RELEASES}/{quote(etiqueta, safe='')}",
            headers=HEADERS_GITHUB,
            timeout=(10, 20),
        )
        if respuesta.status_code == 404:
            continue
        respuesta.raise_for_status()
        release = respuesta.json()

        if release.get("draft") or release.get("prerelease"):
            raise RuntimeError(f"El Release {etiqueta} no es una versión estable publicada.")

        for asset in release.get("assets", []):
            if asset.get("name") != nombre_zip:
                continue

            enlace = asset.get("browser_download_url", "")
            if not enlace.startswith("https://github.com/"):
                raise RuntimeError("El Release devolvió una URL de descarga inesperada.")

            notas = (release.get("body") or "").strip()
            if not notas:
                notas = "Esta versión no incluye una descripción de cambios en GitHub Releases."

            digest = asset.get("digest") or ""
            sha256 = digest.split(":", 1)[1] if digest.startswith("sha256:") else None
            if sha256 and not re.fullmatch(r"[a-fA-F0-9]{64}", sha256):
                raise RuntimeError("El Release contiene un SHA-256 no válido.")

            registrar(f"Release encontrado: {etiqueta}; ZIP: {nombre_zip}")
            return enlace, notas, sha256

        raise RuntimeError(
            f"Encontré el Release {etiqueta}, pero falta el archivo {nombre_zip}. "
            "Sube el ZIP interior de GitHub Actions, no el ZIP exterior."
        )

    raise RuntimeError(
        f"No existe un Release publicado con etiqueta v{version_normalizada} "
        f"o {version_normalizada}. Revisa el tag y la visibilidad del repositorio."
    )


# ============================================================
# DESCARGA EN SEGUNDO PLANO
# ============================================================

class HiloDescarga(QThread):
    progreso = Signal(int)
    terminado = Signal(bool, str)

    def __init__(self, url, destino, sha256_esperado=None):
        super().__init__()
        self.url = url
        self.destino = Path(destino)
        self.sha256_esperado = sha256_esperado

    def run(self):
        parcial = self.destino.with_suffix(".part")
        try:
            total = 0
            recibido = 0
            hash_archivo = hashlib.sha256()

            with requests.get(
                self.url,
                stream=True,
                headers=HEADERS_GITHUB,
                timeout=(15, 90),
                allow_redirects=True,
            ) as respuesta:
                respuesta.raise_for_status()
                total = int(respuesta.headers.get("Content-Length") or 0)
                if not total:
                    self.progreso.emit(-1)  # barra indeterminada

                with parcial.open("wb") as archivo:
                    for bloque in respuesta.iter_content(chunk_size=128 * 1024):
                        if self.isInterruptionRequested():
                            raise InterruptedError("La descarga fue cancelada.")
                        if not bloque:
                            continue
                        archivo.write(bloque)
                        hash_archivo.update(bloque)
                        recibido += len(bloque)
                        if total:
                            self.progreso.emit(min(99, int(100 * recibido / total)))

            if self.isInterruptionRequested():
                raise InterruptedError("La descarga fue cancelada.")
            if total and recibido != total:
                raise RuntimeError("La descarga terminó incompleta.")
            if recibido == 0:
                raise RuntimeError("El archivo descargado está vacío.")
            if self.sha256_esperado and hash_archivo.hexdigest().lower() != self.sha256_esperado.lower():
                raise RuntimeError("El SHA-256 del ZIP no coincide con el publicado en GitHub.")
            if not zipfile.is_zipfile(parcial):
                raise RuntimeError("La descarga no es un ZIP válido.")
            with zipfile.ZipFile(parcial) as archivo_zip:
                malo = archivo_zip.testzip()
                if malo:
                    raise RuntimeError(f"El ZIP contiene un archivo dañado: {malo}")

            parcial.replace(self.destino)
            self.progreso.emit(100)
            self.terminado.emit(True, "")
        except Exception as error:
            parcial.unlink(missing_ok=True)
            registrar(f"Error de descarga: {error}")
            self.terminado.emit(False, str(error))


# ============================================================
# VENTANA CON NOTAS DE VERSIÓN Y BARRA DE PROGRESO
# ============================================================

class VentanaActualizacion(QDialog):
    def __init__(self, version_nueva, notas, url, destino, sha256, parent=None):
        super().__init__(parent)
        self.url = url
        self.destino = destino
        self.sha256 = sha256
        self.descargando = False
        self.error_descarga = ""
        self.hilo = None

        self.setWindowTitle("Actualización de DCO Expedientes")
        self.resize(680, 530)
        self.setMinimumSize(480, 380)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)

        layout = QVBoxLayout(self)

        titulo = QLabel("Actualización disponible")
        titulo.setStyleSheet("font-size: 19px; font-weight: bold;")
        layout.addWidget(titulo)

        version = QLabel(f"Versión instalada: {VERSION_ACTUAL}    →    Nueva versión: {version_nueva}")
        layout.addWidget(version)

        encabezado = QLabel("Novedades de esta versión:")
        encabezado.setStyleSheet("font-weight: bold; margin-top: 9px;")
        layout.addWidget(encabezado)

        self.notas = QTextBrowser()
        self.notas.setOpenExternalLinks(False)
        self.notas.setMarkdown(notas)
        self.notas.setReadOnly(True)
        layout.addWidget(self.notas, 1)

        self.estado = QLabel("¿Quieres descargar e instalar esta actualización?")
        layout.addWidget(self.estado)

        self.barra = QProgressBar()
        self.barra.setRange(0, 100)
        self.barra.setValue(0)
        self.barra.hide()
        layout.addWidget(self.barra)

        botones = QHBoxLayout()
        botones.addStretch()
        self.boton_cancelar = QPushButton("Más tarde")
        self.boton_actualizar = QPushButton("Descargar e instalar")
        self.boton_actualizar.setDefault(True)
        botones.addWidget(self.boton_cancelar)
        botones.addWidget(self.boton_actualizar)
        layout.addLayout(botones)

        self.boton_actualizar.clicked.connect(self.iniciar_descarga)
        self.boton_cancelar.clicked.connect(self.reject)

    def iniciar_descarga(self):
        if self.descargando:
            return
        self.descargando = True
        self.boton_actualizar.setEnabled(False)
        self.boton_cancelar.setText("Cancelar descarga")
        self.barra.show()
        self.estado.setText("Descargando actualización... Puedes leer las novedades mientras tanto.")

        self.hilo = HiloDescarga(self.url, self.destino, self.sha256)
        self.hilo.progreso.connect(self.actualizar_progreso)
        self.hilo.terminado.connect(self.finalizar_descarga)
        self.hilo.start()

    def actualizar_progreso(self, porcentaje):
        if porcentaje == -1:
            self.barra.setRange(0, 0)
            self.estado.setText("Descargando actualización... Calculando tamaño del archivo.")
        else:
            if self.barra.maximum() == 0:
                self.barra.setRange(0, 100)
            self.barra.setValue(porcentaje)
            self.estado.setText(f"Descargando actualización... {porcentaje}%")

    def finalizar_descarga(self, exito, mensaje):
        self.descargando = False
        if self.hilo is not None:
            self.hilo.wait()
        if exito:
            self.accept()
        else:
            self.error_descarga = mensaje
            super().reject()

    def reject(self):
        if self.descargando:
            self.hilo.requestInterruption()
            self.boton_cancelar.setEnabled(False)
            self.estado.setText("Cancelando la descarga...")
            return
        super().reject()

    def closeEvent(self, evento):
        if self.descargando:
            evento.ignore()
            self.reject()
        else:
            super().closeEvent(evento)


# ============================================================
# VALIDACIÓN DE LOS ZIP
# ============================================================

def validar_zip(ruta_zip, sistema):
    with zipfile.ZipFile(ruta_zip) as archivo:
        nombres = archivo.namelist()
        for nombre in nombres:
            ruta_interna = PurePosixPath(nombre)
            if (ruta_interna.is_absolute() or ".." in ruta_interna.parts
                    or "\\" in nombre or not nombre.strip()):
                raise RuntimeError(f"Ruta no permitida dentro del ZIP: {nombre}")

        if sistema == "windows":
            if EXE_WINDOWS not in nombres:
                raise RuntimeError(
                    f"El ZIP de Windows debe contener {EXE_WINDOWS} en la raíz. "
                    "Comprueba que no incluya una carpeta DCO_Expedientes/ adicional "
                    "ni otro ZIP dentro."
                )
        elif not any(nombre.startswith(f"{APP_MAC}/Contents/MacOS/") for nombre in nombres):
            raise RuntimeError(
                f"El ZIP de macOS debe contener {APP_MAC} en la raíz. "
                "Comprueba que no hayas subido el ZIP exterior de GitHub Actions."
            )


# ============================================================
# INSTALACIÓN EXTERNA EN WINDOWS
# ============================================================

def instalar_windows(ruta_zip, exe_actual):
    carpeta_script = Path(tempfile.mkdtemp(prefix="dco_instalador_"))
    script = carpeta_script / "instalar_actualizacion.ps1"
    script.write_text(r'''param(
    [Parameter(Mandatory=$true)][int]$AppProcessId,
    [Parameter(Mandatory=$true)][string]$ZipPath,
    [Parameter(Mandatory=$true)][string]$InstallFolder,
    [Parameter(Mandatory=$true)][string]$ExeName
)
$ErrorActionPreference = "Stop"
$LogPath = Join-Path $HOME "dco_actualizador_instalacion.log"
$Stage = Join-Path $env:TEMP ("dco_stage_" + [guid]::NewGuid().ToString("N"))
try {
    "Iniciando actualización: $(Get-Date)" | Out-File -Append -FilePath $LogPath -Encoding UTF8
    Wait-Process -Id $AppProcessId -Timeout 120 -ErrorAction SilentlyContinue
    if (Get-Process -Id $AppProcessId -ErrorAction SilentlyContinue) {
        throw "La aplicación no terminó de cerrarse."
    }
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $Stage -Force
    $ExeNuevo = Join-Path $Stage $ExeName
    if (-not (Test-Path -LiteralPath $ExeNuevo -PathType Leaf)) {
        throw "El ZIP no contiene el ejecutable en la raíz."
    }
    & robocopy.exe $Stage $InstallFolder /E /R:3 /W:1 /NP /NFL /NDL | Out-Null
    if ($LASTEXITCODE -ge 8) {
        throw "No se pudieron copiar los archivos (robocopy: $LASTEXITCODE)."
    }
    Start-Process -FilePath (Join-Path $InstallFolder $ExeName)
    Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
    "Actualización instalada correctamente: $(Get-Date)" | Out-File -Append -FilePath $LogPath -Encoding UTF8
} catch {
    "ERROR: $($_.Exception.Message)" | Out-File -Append -FilePath $LogPath -Encoding UTF8
} finally {
    if (Test-Path -LiteralPath $Stage) {
        Remove-Item -LiteralPath $Stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}
''', encoding="utf-8")

    subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
         str(script), str(os.getpid()), str(ruta_zip), str(exe_actual.parent), exe_actual.name],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        close_fds=True,
    )


# ============================================================
# INSTALACIÓN EXTERNA EN MACOS
# ============================================================

def instalar_mac(ruta_zip, app_actual):
    carpeta_script = Path(tempfile.mkdtemp(prefix="dco_instalador_"))
    script = carpeta_script / "instalar_actualizacion.sh"
    registro = Path.home() / "dco_actualizador_instalacion.log"

    script.write_text(r'''#!/bin/bash
set -u
PID_ACTUAL="$1"
ZIP="$2"
APP_ACTUAL="$3"
STAGE="$4"
LOG="$5"
exec >> "$LOG" 2>&1
echo "Iniciando actualización: $(date)"

for ((i=0; i<120; i++)); do
    if ! kill -0 "$PID_ACTUAL" 2>/dev/null; then break; fi
    sleep 1
done
if kill -0 "$PID_ACTUAL" 2>/dev/null; then
    echo "ERROR: La aplicación no terminó de cerrarse."
    exit 1
fi

mkdir -p "$STAGE/extraido" || exit 1
if ! /usr/bin/ditto -x -k "$ZIP" "$STAGE/extraido"; then
    echo "ERROR: No se pudo extraer el ZIP."
    exit 1
fi

APP_NUEVA="$STAGE/extraido/DCO_Expedientes.app"
BACKUP="${APP_ACTUAL}.backup"

if [ ! -d "$APP_NUEVA" ]; then
    echo "ERROR: El ZIP no contiene la aplicación esperada."
    exit 1
fi
if [ ! -w "$(dirname "$APP_ACTUAL")" ]; then
    echo "ERROR: No tienes permisos de escritura en la carpeta de instalación."
    exit 1
fi

if [ -e "$BACKUP" ]; then
    echo "ERROR: Ya existe una copia de seguridad; no se sobrescribirá: $BACKUP"
    exit 1
fi
if ! mv "$APP_ACTUAL" "$BACKUP"; then
    echo "ERROR: No se pudo mover la aplicación antigua."
    exit 1
fi
if ! mv "$APP_NUEVA" "$APP_ACTUAL"; then
    echo "ERROR: No se pudo instalar la nueva aplicación. Restaurando..."
    mv "$BACKUP" "$APP_ACTUAL"
    exit 1
fi

if ! open "$APP_ACTUAL"; then
    echo "ERROR: No se pudo iniciar la nueva aplicación. Restaurando..."
    rm -rf "$APP_ACTUAL"
    mv "$BACKUP" "$APP_ACTUAL"
    exit 1
fi
rm -rf "$BACKUP"
rm -f "$ZIP"
echo "Actualización instalada correctamente: $(date)"
rm -rf "$STAGE"
''', encoding="utf-8")
    script.chmod(0o700)

    subprocess.Popen(
        ["/bin/bash", str(script), str(os.getpid()), str(ruta_zip),
         str(app_actual), str(carpeta_script), str(registro)],
        start_new_session=True,
        close_fds=True,
    )


# ============================================================
# CONSULTAS NO BLOQUEANTES Y MONITOR DE ACTUALIZACIONES
# ============================================================

def consultar_actualizacion_disponible():
    """Consulta red fuera del hilo gráfico. None si no hay versión nueva."""
    version_remota = consultar_version_remota()
    if not es_version_nueva(VERSION_ACTUAL, version_remota):
        registrar(f"Sin actualizaciones. Remota: {version_remota}")
        return None

    sistema = plataforma_actual()
    archivo_zip = ZIP_WINDOWS if sistema == "windows" else ZIP_MAC
    url, notas, sha256 = obtener_release_y_archivo(version_remota, archivo_zip)
    return {
        "version": version_remota,
        "url": url,
        "notas": notas,
        "sha256": sha256,
        "sistema": sistema,
    }


class HiloConsulta(QThread):
    """La red se consulta en segundo plano; nunca desde el hilo de Qt."""
    resultado = Signal(object)
    error = Signal(str)

    def run(self):
        try:
            self.resultado.emit(consultar_actualizacion_disponible())
        except Exception as exc:
            registrar(f"Error comprobando actualización: {type(exc).__name__}: {exc}")
            self.error.emit(str(exc))


def descargar_e_instalar(info, parent=None):
    """Muestra las novedades, descarga e inicia el instalador. Retorna True si procede cerrar."""
    destino = None
    try:
        if not getattr(sys, "frozen", False):
            raise RuntimeError(
                "Estás ejecutando app.py desde Python. La actualización automática "
                "solo puede instalarse desde la aplicación compilada (.exe o .app)."
            )

        sistema = info["sistema"]
        if sistema != plataforma_actual():
            raise RuntimeError("La plataforma del Release no coincide con este equipo.")

        # Comprobar ANTES de descargar que realmente es reemplazable.
        instalacion = obtener_instalacion()
        destino = Path(tempfile.gettempdir()) / f"dco_actualizacion_{os.getpid()}.zip"
        ventana = VentanaActualizacion(
            info["version"], info["notas"], info["url"], destino, info["sha256"], parent
        )
        if ventana.exec() != QDialog.DialogCode.Accepted:
            if ventana.error_descarga:
                QMessageBox.warning(
                    parent, "No se descargó la actualización",
                    f"{ventana.error_descarga}\n\nRegistro de errores: {LOG}",
                )
            else:
                registrar(f"Usuario pospuso o canceló versión {info['version']}")
            return False

        validar_zip(destino, sistema)

        # El instalador se inicia, pero NO reemplaza archivos hasta que el proceso actual termine.
        if sistema == "windows":
            instalar_windows(destino, instalacion)
        else:
            instalar_mac(destino, instalacion)

        registrar(f"Instalador externo iniciado para {info['version']}")
        QMessageBox.information(
            parent,
            "Actualización lista",
            "La descarga terminó correctamente.\n\n"
            "DCO Expedientes se cerrará para instalar la actualización y se "
            "volverá a abrir cuando termine si la instalación es correcta.",
        )
        return True
    except Exception as exc:
        registrar(f"ERROR INSTALANDO: {type(exc).__name__}: {exc}")
        QMessageBox.critical(
            parent,
            "Error del actualizador",
            f"No se pudo preparar la actualización:\n\n{exc}\n\nRegistro: {LOG}",
        )
        return False


class MonitorActualizaciones(QObject):
    """Revisa al iniciar y cada 30 minutos, sin interrumpir operaciones activas.

    parent: ventana principal.
    puede_notificar: callback que retorna False si login, búsqueda, PDF o Excel están ocupados.
    """

    def __init__(self, parent, puede_notificar, intervalo_ms=INTERVALO_REVISION_MS):
        super().__init__(parent)
        self.ventana = parent
        self.puede_notificar = puede_notificar
        self.hilo = None
        self.pendiente = None
        self.ultima_notificada = None
        self.mostrando_dialogo = False
        self.consulta_manual = False
        self.deteniendo = False

        self.temporizador = QTimer(self)
        self.temporizador.setInterval(intervalo_ms)
        self.temporizador.timeout.connect(self.comprobar)

        # Al terminar una operación se mostrará el aviso pendiente (si existe).
        # Este temporizador NO hace peticiones a GitHub.
        self.temporizador_pendientes = QTimer(self)
        self.temporizador_pendientes.setInterval(5000)
        self.temporizador_pendientes.timeout.connect(self.mostrar_si_es_posible)

    def iniciar(self):
        if self.deteniendo:
            return
        self.temporizador.start()
        self.temporizador_pendientes.start()
        self.comprobar()

    def detener(self):
        """Evita que se destruya un QThread activo durante el cierre de la app."""
        self.deteniendo = True
        self.temporizador.stop()
        self.temporizador_pendientes.stop()
        if self.hilo is not None and self.hilo.isRunning():
            self.hilo.wait()  # Las solicitudes de red tienen timeout; esperar cierre seguro.

    def comprobar(self, manual=False):
        if self.deteniendo or self.mostrando_dialogo:
            return
        if self.hilo is not None and self.hilo.isRunning():
            if manual:
                QMessageBox.information(
                    self.ventana, "Actualizaciones", "Ya se está comprobando GitHub."
                )
            return
        if self.pendiente is not None:
            if manual:
                self.ultima_notificada = None  # Permite reabrir aviso manualmente.
                self.mostrar_si_es_posible()
            return

        if not getattr(sys, "frozen", False):
            if manual:
                QMessageBox.information(
                    self.ventana, "Modo desarrollo",
                    "La instalación automática solo está disponible en la aplicación compilada.",
                )
            return

        self.consulta_manual = bool(manual)
        self.hilo = HiloConsulta(self)
        self.hilo.resultado.connect(self.consulta_finalizada)
        self.hilo.error.connect(self.consulta_fallida)
        self.hilo.start()

    def consulta_finalizada(self, info):
        if self.deteniendo:
            return
        manual = self.consulta_manual
        self.consulta_manual = False
        if info is None:
            if manual:
                QMessageBox.information(
                    self.ventana, "DCO Expedientes actualizado",
                    f"Ya tienes la última versión ({VERSION_ACTUAL}).",
                )
            return
        if not manual and info["version"] == self.ultima_notificada:
            return
        registrar(f"Nueva versión disponible: {info['version']}")
        self.pendiente = info
        if manual:
            self.ultima_notificada = None
        self.mostrar_si_es_posible()

    def consulta_fallida(self, mensaje):
        if self.deteniendo:
            return
        manual = self.consulta_manual
        self.consulta_manual = False
        if manual:
            QMessageBox.warning(
                self.ventana,
                "No se pudieron comprobar las actualizaciones",
                f"{mensaje}\n\nRegistro: {LOG}",
            )

    def mostrar_si_es_posible(self):
        if self.deteniendo or self.mostrando_dialogo or not self.pendiente:
            return
        if self.pendiente["version"] == self.ultima_notificada:
            self.pendiente = None
            return
        if not self.puede_notificar():
            return
        if QApplication.activeModalWidget() is not None:
            return
        info = self.pendiente
        self.pendiente = None
        self.ultima_notificada = info["version"]
        self.mostrando_dialogo = True
        try:
            debe_cerrar = descargar_e_instalar(info, self.ventana)
        finally:
            self.mostrando_dialogo = False
        if debe_cerrar:
            self.detener()
            QApplication.instance().quit()


# ============================================================
# FUNCIÓN LEGADA (COMPROBACIÓN MANUAL SIN MONITOR)
# ============================================================

def verificar_e_instalar_actualizacion(parent=None):
    """Compatibilidad con la llamada antigua. Bloquea la UI mientras consulta GitHub."""
    try:
        info = consultar_actualizacion_disponible()
        if info is None:
            return False
        return descargar_e_instalar(info, parent)
    except Exception as exc:
        registrar(f"Error de comprobación inicial: {exc}")
        return False