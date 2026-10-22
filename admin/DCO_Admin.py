"""Panel privado DCO Feedback. Ejecutar: python admin/DCO_Admin.py."""
import csv
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from PySide6.QtWidgets import (
    QApplication, QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox,
    QPlainTextEdit, QInputDialog, QFileDialog, QSplitter, QGroupBox,
)
from PySide6.QtCore import Qt
from feedback_api import FeedbackAPI, FeedbackError
from feedback_config import configurado


class LoginDialog(QDialog):
    def __init__(self, api):
        super().__init__()
        self.api = api
        self.setWindowTitle("DCO Feedback — Acceso de administrador")
        self.resize(390, 210)
        caja = QVBoxLayout(self)
        caja.addWidget(QLabel("Inicia sesión con tu cuenta administrativa de Supabase Auth."))
        self.correo = QLineEdit()
        self.correo.setPlaceholderText("Correo electrónico")
        caja.addWidget(self.correo)
        self.clave = QLineEdit()
        self.clave.setEchoMode(QLineEdit.EchoMode.Password)
        self.clave.setPlaceholderText("Contraseña")
        caja.addWidget(self.clave)
        entrar = QPushButton("Iniciar sesión")
        entrar.clicked.connect(self.ingresar)
        caja.addWidget(entrar)

    def ingresar(self):
        try:
            self.api.login_admin(self.correo.text(), self.clave.text())
            self.clave.clear()
            self.accept()
        except Exception as exc:
            QMessageBox.warning(self, "No autorizado", str(exc))


class AdminApp(QWidget):
    def __init__(self, api):
        super().__init__()
        self.api = api
        self.campañas = []
        self.respuestas = []
        self.sugerencias = []
        self.setWindowTitle("DCO Feedback — Administración")
        self.resize(1040, 690)
        base = QVBoxLayout(self)
        base.addWidget(QLabel("ADMINISTRACIÓN DE ENCUESTAS · DCO EXPEDIENTES"))
        base.addWidget(QLabel("Las campañas se muestran a los compañeros en un plazo aproximado de 5 minutos."))
        splitter = QSplitter(Qt.Orientation.Horizontal)
        base.addWidget(splitter)
        izquierda = QWidget()
        caja = QVBoxLayout(izquierda)
        caja.addWidget(QLabel("Crear campaña"))
        self.titulo = QLineEdit()
        self.titulo.setPlaceholderText("Título de la encuesta")
        caja.addWidget(self.titulo)
        self.descripcion = QPlainTextEdit()
        self.descripcion.setPlaceholderText("Explica qué te gustaría evaluar...")
        self.descripcion.setMaximumHeight(110)
        caja.addWidget(self.descripcion)
        crear = QPushButton("Crear campaña (inactiva)")
        crear.clicked.connect(self.crear)
        caja.addWidget(crear)
        caja.addWidget(QLabel("Campañas"))
        self.tabla_campañas = QTableWidget(0, 3)
        self.tabla_campañas.setHorizontalHeaderLabels(["Nombre", "Estado", "Creada"])
        self.tabla_campañas.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla_campañas.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla_campañas.itemSelectionChanged.connect(self.mostrar_respuestas)
        caja.addWidget(self.tabla_campañas)
        acciones = QHBoxLayout()
        activar = QPushButton("Activar")
        activar.clicked.connect(self.activar)
        parar = QPushButton("Detener")
        parar.clicked.connect(self.detener)
        acciones.addWidget(activar)
        acciones.addWidget(parar)
        caja.addLayout(acciones)
        actualizar = QPushButton("Actualizar resultados")
        actualizar.clicked.connect(self.recargar)
        caja.addWidget(actualizar)
        splitter.addWidget(izquierda)

        derecha = QWidget()
        caja2 = QVBoxLayout(derecha)
        self.resumen = QLabel("Selecciona una campaña.")
        caja2.addWidget(self.resumen)
        self.tabla_respuestas = QTableWidget(0, 5)
        self.tabla_respuestas.setHorizontalHeaderLabels(["Fecha", "★", "Le gusta", "Mejoraría", "Área"])
        self.tabla_respuestas.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        caja2.addWidget(self.tabla_respuestas)
        btn_csv = QPushButton("Exportar respuestas de esta campaña a CSV")
        btn_csv.clicked.connect(self.exportar_respuestas)
        caja2.addWidget(btn_csv)
        caja2.addWidget(QLabel("Sugerencias enviadas sin encuesta"))
        self.tabla_sugerencias = QTableWidget(0, 2)
        self.tabla_sugerencias.setHorizontalHeaderLabels(["Fecha", "Sugerencia"])
        self.tabla_sugerencias.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        caja2.addWidget(self.tabla_sugerencias)
        btn_sugerencias = QPushButton("Exportar sugerencias a CSV")
        btn_sugerencias.clicked.connect(self.exportar_sugerencias)
        caja2.addWidget(btn_sugerencias)
        splitter.addWidget(derecha)
        splitter.setSizes([400, 640])
        self.recargar()

    def _error(self, exc):
        QMessageBox.critical(self, "DCO Feedback", str(exc))

    def campaña_actual(self):
        fila = self.tabla_campañas.currentRow()
        return self.campañas[fila] if 0 <= fila < len(self.campañas) else None

    def recargar(self):
        try:
            self.campañas = self.api.listar_campañas()
            self.sugerencias = self.api.listar_sugerencias()
        except Exception as exc:
            self._error(exc)
            return
        self.tabla_campañas.setRowCount(len(self.campañas))
        for i, c in enumerate(self.campañas):
            for j, value in enumerate((c["title"], "ACTIVA" if c["active"] else "Inactiva", c["created_at"][:16])):
                self.tabla_campañas.setItem(i, j, QTableWidgetItem(str(value)))
        self.tabla_campañas.resizeColumnsToContents()
        self.tabla_sugerencias.setRowCount(len(self.sugerencias))
        for i, s in enumerate(self.sugerencias):
            self.tabla_sugerencias.setItem(i, 0, QTableWidgetItem(s["created_at"][:16]))
            self.tabla_sugerencias.setItem(i, 1, QTableWidgetItem(s["body"]))
        self.tabla_sugerencias.resizeColumnsToContents()
        if self.campañas:
            self.tabla_campañas.selectRow(0)
        else:
            self.resumen.setText("No hay campañas. Crea una y después actívala.")
            self.tabla_respuestas.setRowCount(0)

    def crear(self):
        try:
            self.api.crear_campaña(self.titulo.text(), self.descripcion.toPlainText())
            self.titulo.clear()
            self.descripcion.clear()
            self.recargar()
            QMessageBox.information(self, "Campaña", "Campaña creada. Pulsa Activar cuando quieras publicarla.")
        except Exception as exc:
            self._error(exc)

    def activar(self):
        c = self.campaña_actual()
        if c is None:
            return
        if QMessageBox.question(self, "Activar encuesta", f"¿Mostrar la encuesta '{c['title']}' a todos los equipos?") != QMessageBox.StandardButton.Yes:
            return
        try:
            self.api.activar_campaña(c["id"])
            self.recargar()
        except Exception as exc:
            self._error(exc)

    def detener(self):
        try:
            self.api.detener_campaña()
            self.recargar()
        except Exception as exc:
            self._error(exc)

    def mostrar_respuestas(self):
        c = self.campaña_actual()
        if c is None:
            return
        try:
            self.respuestas = self.api.listar_respuestas(c["id"])
        except Exception as exc:
            self._error(exc)
            return
        self.tabla_respuestas.setRowCount(len(self.respuestas))
        for i, r in enumerate(self.respuestas):
            datos = (r["created_at"][:16], r["rating"], r["liked"], r["improve"], r["category"])
            for j, valor in enumerate(datos):
                self.tabla_respuestas.setItem(i, j, QTableWidgetItem(str(valor)))
        cantidad = len(self.respuestas)
        promedio = round(sum(r["rating"] for r in self.respuestas) / cantidad, 2) if cantidad else 0
        self.resumen.setText(f"{c['title']}  ·  Respuestas: {cantidad}  ·  Promedio: {promedio}/5")
        self.tabla_respuestas.resizeColumnsToContents()

    def _guardar_csv(self, nombre, filas, columnas):
        if not filas:
            QMessageBox.information(self, "Exportación", "No hay registros para exportar.")
            return
        archivo, _ = QFileDialog.getSaveFileName(self, "Guardar CSV", nombre, "CSV (*.csv)")
        if not archivo:
            return
        try:
            with open(archivo, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=columnas, extrasaction="ignore")
                writer.writeheader()
                for fila in filas:
                    # Evita que Excel interprete comentarios como fórmulas CSV.
                    seguro = {}
                    for clave in columnas:
                        valor = fila.get(clave, "")
                        if isinstance(valor, str) and valor.lstrip().startswith(("=", "+", "-", "@")):
                            valor = "'" + valor
                        seguro[clave] = valor
                    writer.writerow(seguro)
            QMessageBox.information(self, "Archivo guardado", "Se guardó el CSV correctamente.")
        except OSError as exc:
            self._error(exc)

    def exportar_respuestas(self):
        self._guardar_csv("dco_respuestas.csv", self.respuestas,
                          ["created_at", "rating", "liked", "improve", "category"])

    def exportar_sugerencias(self):
        self._guardar_csv("dco_sugerencias.csv", self.sugerencias, ["created_at", "body"])


if __name__ == "__main__":
    app = QApplication(sys.argv)
    if not configurado():
        QMessageBox.critical(None, "Sin configurar", "Edita feedback_config.py con la URL y la clave publicable de Supabase.")
        sys.exit(1)
    api = FeedbackAPI(admin=True)
    if LoginDialog(api).exec() == QDialog.DialogCode.Accepted:
        ventana = AdminApp(api)
        ventana.show()
        sys.exit(app.exec())
