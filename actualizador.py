import os
import sys
import subprocess
import zipfile
import requests
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import QProgressDialog, QMessageBox

# Versión actual de tu aplicación
VERSION_ACTUAL = "1.0.0"
VERSION_URL = "https://raw.githubusercontent.com/jomancrazy/DCO_LICENCIAS/main/version.txt"

# URLs de descarga directa desde tus GitHub Releases
URL_WINDOWS = "https://github.com/jomancrazy/DCO_LICENCIAS/releases/latest/download/DCO_Expedientes_Windows.zip"
URL_MAC = "https://github.com/jomancrazy/DCO_LICENCIAS/releases/latest/download/DCO_Expedientes_Mac.zip"

class DescargaThread(QThread):
    progreso = Signal(int)
    terminado = Signal(bool, str)

    def __init__(self, url_descarga, ruta_destino):
        super().__init__()
        self.url_descarga = url_descarga
        self.ruta_destino = ruta_destino

    def run(self):
        try:
            response = requests.get(self.url_descarga, stream=True, timeout=30)
            if response.status_code != 200:
                self.terminado.emit(False, "No se pudo conectar al servidor de descargas.")
                return

            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(self.ruta_destino, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            percent = int((downloaded / total_size) * 100)
                            self.progreso.emit(percent)
                            
            self.terminado.emit(True, "")
        except Exception as e:
            self.terminado.emit(False, str(e))

def verificar_e_instalar_actualizacion(parent=None):
    """Verifica si hay actualización al iniciar y gestiona todo el flujo."""
    try:
        respuesta = requests.get(VERSION_URL, timeout=5)
        if respuesta.status_code != 200:
            return False
            
        version_remota = respuesta.text.strip()
        if version_remota == VERSION_ACTUAL:
            return False # No hay actualización, la app abre normal

        # Preguntar al usuario si desea actualizar
        msg = QMessageBox(parent)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Actualización disponible")
        msg.setText(f"¡Hay una nueva versión disponible ({version_remota})!")
        msg.setInformativeText("La aplicación se actualizará y se reiniciará automáticamente. ¿Deseas continuar?")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return False

        # Configurar rutas y descarga
        url_descarga = URL_WINDOWS if sys.platform.startswith("win") else URL_MAC
        temp_dir = os.path.expanduser("~")
        zip_path = os.path.join(temp_dir, "dco_update.zip")
        app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

        # Diálogo de progreso
        progress = QProgressDialog("Descargando actualización...", "Cancelar", 0, 100, parent)
        progress.setWindowTitle("Actualizando DCO Expedientes")
        progress.setWindowModality(True)
        progress.setAutoClose(True)
        progress.show()

        thread = DescargaThread(url_descarga, zip_path)
        
        def actualizar_barra(val):
            progress.setValue(val)
            if progress.wasCanceled():
                thread.terminate()

        thread.progreso.connect(actualizar_barra)
        
        resultado_descarga = {"exito": False, "error": ""}
        
        def al_terminar(exito, error):
            resultado_descarga["exito"] = exito
            resultado_descarga["error"] = error
            progress.close()

        thread.terminado.connect(al_terminar)
        thread.start()
        
        while thread.isRunning():
            from PySide6.QtCore import QCoreApplication
            QCoreApplication.processEvents()

        if not resultado_descarga["exito"]:
            QMessageBox.critical(parent, "Error", f"No se pudo completar la descarga: {resultado_descarga['error']}")
            return False

        # Descomprimir y reemplazar archivos
        progress_extract = QProgressDialog("Instalando actualización...", "", 0, 0, parent)
        progress_extract.setWindowModality(True)
        progress_extract.show()
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(app_dir)
            
        os.remove(zip_path)

        # Si es Mac, quitar la cuarentena de Apple automáticamente
        if sys.platform == "darwin":
            app_path = os.path.join(app_dir, "DCO_Expedientes.app")
            subprocess.run(["xattr", "-cr", app_path], capture_output=True)

        progress_extract.close()

        QMessageBox.information(parent, "Actualización exitosa", "La aplicación se reiniciará para aplicar los cambios.")
        
        # Reiniciar la aplicación automáticamente
        if sys.platform.startswith("win"):
            subprocess.Popen([sys.executable] + sys.argv)
        else:
            os.execv(sys.executable, [sys.executable] + sys.argv)
            
        sys.exit(0)
        
    except Exception as e:
        print(f"Error en actualizador: {e}")
        return False