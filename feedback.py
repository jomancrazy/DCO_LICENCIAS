"""Ventanas y monitor no intrusivo de DCO Feedback para PySide6."""
import threading
from PySide6.QtCore import QObject, QTimer, Signal, Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextBrowser,
    QTextEdit, QComboBox, QMessageBox, QApplication,
)

from feedback_config import configurado, POLL_INTERVAL_MS
from feedback_api import FeedbackAPI, FeedbackError


CATEGORIAS = (
    "Facilidad de uso", "Velocidad", "Exportación a Excel",
    "Búsquedas", "Nuevas funciones", "Otro",
)


class EncuestaDialog(QDialog):
    envio_terminado = Signal(bool, str)

    def __init__(self, campaña, parent=None):
        super().__init__(parent)
        self.campaña = campaña
        self.calificacion = 0
        self.enviando = False
        self.setWindowTitle("DCO Feedback — Tu opinión")
        self.resize(570, 560)
        self.setMinimumWidth(460)
        layout = QVBoxLayout(self)
        subtitulo = QLabel("Encuesta voluntaria · No solicita tu nombre ni datos de expedientes")
        subtitulo.setWordWrap(True)
        layout.addWidget(subtitulo)
        titulo = QLabel(str(campaña.get("title") or "Queremos conocer tu opinión"))
        titulo.setStyleSheet("font-size: 18px; font-weight: bold;")
        titulo.setWordWrap(True)
        layout.addWidget(titulo)
        descripcion = QTextBrowser()
        descripcion.setMaximumHeight(108)
        descripcion.setMarkdown(str(campaña.get("description") or "Ayúdanos a mejorar DCO Expedientes."))
        layout.addWidget(descripcion)
        layout.addWidget(QLabel("¿Cómo calificas la aplicación?"))
        estrellas = QHBoxLayout()
        self.botones_estrellas = []
        for numero in range(1, 6):
            boton = QPushButton("☆")
            boton.setMinimumHeight(46)
            boton.setStyleSheet("font-size: 25px;")
            boton.clicked.connect(lambda _, n=numero: self.seleccionar(n))
            estrellas.addWidget(boton)
            self.botones_estrellas.append(boton)
        layout.addLayout(estrellas)
        layout.addWidget(QLabel("¿Qué te gusta de DCO Expedientes? (opcional)"))
        self.gusta = QTextEdit()
        self.gusta.setPlaceholderText("Cuéntanos qué funciones te ayudan más...")
        self.gusta.setMaximumHeight(75)
        layout.addWidget(self.gusta)
        layout.addWidget(QLabel("¿Qué no te gusta o cambiarías? (opcional)"))
        self.mejorar = QTextEdit()
        self.mejorar.setPlaceholderText("Ideas, problemas o mejoras que te gustaría ver...")
        self.mejorar.setMaximumHeight(75)
        layout.addWidget(self.mejorar)
        layout.addWidget(QLabel("¿Qué área priorizarías?"))
        self.categoria = QComboBox()
        self.categoria.addItems(CATEGORIAS)
        layout.addWidget(self.categoria)
        self.estado = QLabel("")
        self.estado.setWordWrap(True)
        layout.addWidget(self.estado)
        acciones = QHBoxLayout()
        self.mas_tarde = QPushButton("Más tarde")
        self.mas_tarde.clicked.connect(self.reject)
        self.enviar = QPushButton("Enviar opinión")
        self.enviar.clicked.connect(self.enviar_respuesta)
        acciones.addWidget(self.mas_tarde)
        acciones.addWidget(self.enviar)
        layout.addLayout(acciones)
        self.envio_terminado.connect(self._fin_envio)

    def seleccionar(self, n):
        self.calificacion = n
        for indice, boton in enumerate(self.botones_estrellas, 1):
            boton.setText("★" if indice <= n else "☆")

    def enviar_respuesta(self):
        if not self.calificacion:
            QMessageBox.information(self, "Calificación", "Selecciona entre 1 y 5 estrellas.")
            return
        if self.enviando:
            return
        self.enviando = True
        self.enviar.setEnabled(False)
        self.mas_tarde.setEnabled(False)
        self.estado.setText("Enviando opinión...")
        datos = (
            self.campaña["id"], self.calificacion,
            self.gusta.toPlainText(), self.mejorar.toPlainText(),
            self.categoria.currentText(),
        )

        def tarea():
            try:
                FeedbackAPI().enviar_respuesta(*datos)
                self.envio_terminado.emit(True, "¡Gracias por ayudarnos a mejorar DCO!")
            except Exception as exc:
                self.envio_terminado.emit(False, str(exc))
        threading.Thread(target=tarea, daemon=True).start()

    def _fin_envio(self, ok, mensaje):
        self.enviando = False
        self.enviar.setEnabled(True)
        self.mas_tarde.setEnabled(True)
        self.estado.setText(mensaje)
        if ok:
            QMessageBox.information(self, "DCO Feedback", mensaje)
            self.accept()
        else:
            QMessageBox.warning(self, "No se pudo enviar", mensaje)

    def reject(self):
        if not self.enviando:
            super().reject()


class SugerenciaDialog(QDialog):
    envio_terminado = Signal(bool, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enviar una sugerencia — DCO")
        self.resize(490, 310)
        self.enviando = False
        layout = QVBoxLayout(self)
        etiqueta = QLabel("¿Qué cambiarías o añadirías en DCO Expedientes?\n"
                          "No escribas datos personales ni información de expedientes.")
        etiqueta.setWordWrap(True)
        layout.addWidget(etiqueta)
        self.texto = QTextEdit()
        self.texto.setPlaceholderText("Describe tu idea o un problema que hayas encontrado...")
        layout.addWidget(self.texto)
        self.estado = QLabel("")
        layout.addWidget(self.estado)
        acciones = QHBoxLayout()
        self.cancelar = QPushButton("Cancelar")
        self.cancelar.clicked.connect(self.reject)
        self.enviar = QPushButton("Enviar sugerencia")
        self.enviar.clicked.connect(self.enviar_datos)
        acciones.addWidget(self.cancelar)
        acciones.addWidget(self.enviar)
        layout.addLayout(acciones)
        self.envio_terminado.connect(self._fin_envio)

    def enviar_datos(self):
        texto = self.texto.toPlainText().strip()
        if len(texto) < 5:
            QMessageBox.warning(self, "Sugerencia", "Escribe una sugerencia de al menos 5 caracteres.")
            return
        self.enviando = True
        self.cancelar.setEnabled(False)
        self.enviar.setEnabled(False)
        self.estado.setText("Enviando...")

        def tarea():
            try:
                FeedbackAPI().enviar_sugerencia(texto)
                self.envio_terminado.emit(True, "¡Gracias! Tu sugerencia fue enviada.")
            except Exception as exc:
                self.envio_terminado.emit(False, str(exc))
        threading.Thread(target=tarea, daemon=True).start()

    def _fin_envio(self, ok, mensaje):
        self.enviando = False
        self.cancelar.setEnabled(True)
        self.enviar.setEnabled(True)
        self.estado.setText(mensaje)
        if ok:
            QMessageBox.information(self, "DCO Feedback", mensaje)
            self.accept()
        else:
            QMessageBox.warning(self, "No se pudo enviar", mensaje)

    def reject(self):
        if not self.enviando:
            super().reject()


class MonitorFeedback(QObject):
    consulta_terminada = Signal(object, str)

    def __init__(self, parent, puede_mostrar):
        super().__init__(parent)
        self.ventana = parent
        self.puede_mostrar = puede_mostrar
        self.consultando = False
        self.campaña_pendiente = None
        self.pospuestas_en_sesion = set()
        self.dialogo_abierto = False
        self.timer = QTimer(self)
        self.timer.setInterval(POLL_INTERVAL_MS)
        self.timer.timeout.connect(self.comprobar)
        self.pendiente_timer = QTimer(self)
        self.pendiente_timer.setInterval(15000)
        self.pendiente_timer.timeout.connect(self.mostrar_si_es_posible)
        self.consulta_terminada.connect(self._recibir_resultado)

    def iniciar(self):
        if configurado():
            self.timer.start()
            self.pendiente_timer.start()
            QTimer.singleShot(16000, self.comprobar)

    def detener(self):
        self.timer.stop()
        self.pendiente_timer.stop()

    def comprobar(self):
        if not configurado() or self.consultando or self.dialogo_abierto:
            return
        self.consultando = True

        def tarea():
            try:
                campaña = FeedbackAPI().campaña_activa()
                self.consulta_terminada.emit(campaña, "")
            except Exception as exc:
                self.consulta_terminada.emit(None, str(exc))
        threading.Thread(target=tarea, daemon=True).start()

    def _recibir_resultado(self, campaña, error):
        self.consultando = False
        if error:
            # Un fallo en el servicio opcional de encuestas no bloquea DCO.
            return
        if campaña is None:
            self.campaña_pendiente = None
            return
        if campaña["id"] not in self.pospuestas_en_sesion:
            self.campaña_pendiente = campaña
        self.mostrar_si_es_posible()

    def mostrar_si_es_posible(self):
        if not self.campaña_pendiente or self.dialogo_abierto:
            return
        if not self.puede_mostrar() or QApplication.activeModalWidget() is not None:
            return
        campaña = self.campaña_pendiente
        self.campaña_pendiente = None
        self.dialogo_abierto = True
        dialogo = EncuestaDialog(campaña, self.ventana)
        try:
            if dialogo.exec() != QDialog.DialogCode.Accepted:
                self.pospuestas_en_sesion.add(campaña["id"])
        finally:
            self.dialogo_abierto = False

    def enviar_sugerencia_manual(self):
        if not configurado():
            QMessageBox.information(
                self.ventana, "DCO Feedback",
                "La función todavía requiere configurar el servidor en feedback_config.py.",
            )
            return
        if not self.puede_mostrar():
            QMessageBox.information(
                self.ventana, "DCO Feedback",
                "Finaliza la operación actual antes de enviar una sugerencia.",
            )
            return
        SugerenciaDialog(self.ventana).exec()
