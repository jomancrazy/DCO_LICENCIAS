from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout,
)
from PySide6.QtCore import Qt, QTimer


META_LUNES_JUEVES = 5
META_VIERNES = 10
DIAS = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")


def meta_del_dia(fecha: date) -> int:
    """5 de lunes a jueves; 10 viernes; 0 sábados y domingos."""
    dia = fecha.weekday()
    if dia == 4:
        return META_VIERNES
    return META_LUNES_JUEVES if dia < 4 else 0


@dataclass(frozen=True)
class Resumen:
    fecha: date
    exportados: int
    meta: int
    porcentaje: int
    icono: str
    estado: str


def evaluar_dia(fecha: date, exportados: int) -> Resumen:
    """Indicador de AVANCE contra la meta, no satisfacción del trabajador."""
    meta = meta_del_dia(fecha)
    if not meta:
        return Resumen(fecha, exportados, meta, 0, "☕", "Día sin meta")
    porcentaje = round(100 * exportados / meta)
    if exportados >= meta:
        return Resumen(fecha, exportados, meta, porcentaje, "😄", "¡Meta alcanzada!")
    if exportados >= meta * 0.6:
        return Resumen(fecha, exportados, meta, porcentaje, "🙂", "Buen avance")
    return Resumen(fecha, exportados, meta, porcentaje, "😐", "En progreso")


class LectorMetricas:
    """Solo lectura de la BD del historial; nunca modifica datos de Excel."""

    def __init__(self, ruta_sqlite: str | Path):
        self.ruta_sqlite = Path(ruta_sqlite)

    def conteos(self, inicio: date, fin: date) -> dict[str, int]:
        if fin < inicio:
            return {}
        # La fecha se guarda como ISO local, p.ej. 2026-10-21T14:20:00-06:00.
        # Los límites de fecha hacen que cada exportación se atribuya al día
        # de trabajo en que se registró, no al día UTC del servidor.
        if not self.ruta_sqlite.is_file():
            return {}
        with sqlite3.connect(str(self.ruta_sqlite), timeout=10) as conexion:
            filas = conexion.execute(
                """SELECT substr(fecha, 1, 10), COUNT(DISTINCT UPPER(TRIM(expediente)))
                   FROM exportaciones
                   WHERE fecha >= ? AND fecha < ?
                     AND expediente IS NOT NULL AND TRIM(expediente) <> ''
                   GROUP BY substr(fecha, 1, 10)""",
                (inicio.isoformat(), (fin + timedelta(days=1)).isoformat()),
            ).fetchall()
        return {dia: cantidad for dia, cantidad in filas}

    def resumen(self, fecha: date | None = None) -> Resumen:
        fecha = fecha or date.today()
        cantidad = self.conteos(fecha, fecha).get(fecha.isoformat(), 0)
        return evaluar_dia(fecha, cantidad)

    def resumen_semanal(self, fecha: date | None = None) -> dict:
        fecha = fecha or date.today()
        lunes = fecha - timedelta(days=fecha.weekday())
        domingo = lunes + timedelta(days=6)
        conteos = self.conteos(lunes, domingo)
        filas = []
        for i in range(7):
            dia = lunes + timedelta(days=i)
            filas.append(evaluar_dia(dia, conteos.get(dia.isoformat(), 0)))
        # "Hasta hoy" compara solo las metas que ya corresponden, no
        # compara un lunes contra las 30 unidades de toda la semana.
        meta_hasta_hoy = sum(item.meta for item in filas if item.fecha <= fecha)
        avance_hasta_hoy = sum(item.exportados for item in filas if item.fecha <= fecha)
        return {
            "dias": filas,
            "exportados_semana": sum(item.exportados for item in filas),
            "meta_semanal": sum(item.meta for item in filas),
            "meta_hasta_hoy": meta_hasta_hoy,
            "avance_hasta_hoy": avance_hasta_hoy,
        }


class VentanaMetricas(QDialog):
    """Detalle de la semana actual; no expone datos personales de expedientes."""

    def __init__(self, lector: LectorMetricas, parent=None):
        super().__init__(parent)
        self.lector = lector
        self.setWindowTitle("Métricas de exportación · DCO")
        self.resize(580, 410)
        contenedor = QVBoxLayout(self)
        titulo = QLabel("RESUMEN DE LA SEMANA")
        titulo.setStyleSheet("font-size: 17px; font-weight: 700;")
        contenedor.addWidget(titulo)
        semana = self.lector.resumen_semanal()
        self.subtitulo = QLabel(
            f"Hasta hoy: {semana['avance_hasta_hoy']} / "
            f"{semana['meta_hasta_hoy']}  ·  "
            f"Meta completa de la semana: {semana['meta_semanal']}"
        )
        self.subtitulo.setWordWrap(True)
        contenedor.addWidget(self.subtitulo)
        tabla = QTableWidget(7, 4, self)
        tabla.setHorizontalHeaderLabels(["Día", "Exportados", "Meta", "Avance"])
        tabla.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        tabla.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tabla.verticalHeader().setVisible(False)
        tabla.horizontalHeader().setStretchLastSection(True)
        for fila, item in enumerate(semana["dias"]):
            columnas = (
                f"{DIAS[item.fecha.weekday()].capitalize()} {item.fecha:%d/%m}",
                str(item.exportados),
                str(item.meta) if item.meta else "—",
                (f"{item.icono}  {item.porcentaje}%" if item.meta else "Sin meta"),
            )
            for col, texto in enumerate(columnas):
                tabla.setItem(fila, col, QTableWidgetItem(texto))
        tabla.resizeColumnsToContents()
        contenedor.addWidget(tabla)
        nota = QLabel(
            "Se cuentan expedientes distintos exportados correctamente por día "
            "en esta computadora. No son licencias oficialmente emitidas ni "
            "incluye registros hechos en otros equipos."
        )
        nota.setWordWrap(True)
        nota.setStyleSheet("color: palette(mid);")
        contenedor.addWidget(nota)
        btn = QPushButton("Cerrar")
        btn.clicked.connect(self.accept)
        contenedor.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)


class PanelMetricas(QFrame):
    """Tarjeta compacta para la pantalla principal. Actualizable con refrescar()."""

    def __init__(self, ruta_sqlite: str | Path, parent=None):
        super().__init__(parent)
        self.lector = LectorMetricas(ruta_sqlite)
        self.setObjectName("TarjetaMetricas")
        self.setStyleSheet("""
            QFrame#TarjetaMetricas {
                background-color: palette(alternate-base);
                border: 1px solid palette(mid);
                border-radius: 9px;
            }
            QFrame#TarjetaMetricas QLabel { border: none; background: transparent; }
        """)
        principal = QVBoxLayout(self)
        principal.setContentsMargins(14, 10, 14, 10)
        principal.setSpacing(5)
        linea = QHBoxLayout()
        self.carita = QLabel("😐")
        self.carita.setStyleSheet("font-size: 33px; background: transparent; border: none;")
        linea.addWidget(self.carita)
        texto = QVBoxLayout()
        titulo = QLabel("META DIARIA · EXPEDIENTES EXPORTADOS")
        titulo.setStyleSheet("font-size: 11px; font-weight: 700;")
        texto.addWidget(titulo)
        self.cantidad = QLabel("0 / 5")
        self.cantidad.setStyleSheet("font-size: 21px; font-weight: 700;")
        texto.addWidget(self.cantidad)
        self.estado = QLabel("En progreso")
        texto.addWidget(self.estado)
        linea.addLayout(texto, 1)
        btn = QPushButton("Ver semana")
        btn.clicked.connect(self.ver_semana)
        linea.addWidget(btn)
        principal.addLayout(linea)
        self.barra = QProgressBar()
        self.barra.setRange(0, 100)
        self.barra.setTextVisible(True)
        principal.addWidget(self.barra)
        self.semana = QLabel("")
        self.semana.setStyleSheet("font-size: 11px; color: palette(text);")
        principal.addWidget(self.semana)
        self.ultima_fecha = None
        self.refrescar()
        # Refresca también al cambiar el día sin exigir reiniciar la aplicación.
        self.temporizador = QTimer(self)
        self.temporizador.setInterval(60 * 1000)
        self.temporizador.timeout.connect(self.refrescar)
        self.temporizador.start()

    def refrescar(self):
        try:
            hoy = date.today()
            datos = self.lector.resumen(hoy)
            semana = self.lector.resumen_semanal(hoy)
            self.carita.setText(datos.icono)
            self.estado.setText(datos.estado)
            self.cantidad.setText(
                f"{datos.exportados} / {datos.meta}"
                if datos.meta else f"{datos.exportados} · Sin meta hoy"
            )
            self.barra.setValue(min(100, datos.porcentaje))
            self.barra.setFormat(
                f"{datos.porcentaje}%" if datos.meta else "Sin meta"
            )
            self.semana.setText(
                f"Semana: {semana['avance_hasta_hoy']} / "
                f"{semana['meta_hasta_hoy']} hasta hoy "
                f"· meta semanal: {semana['meta_semanal']}"
            )
            self.ultima_fecha = hoy
        except (sqlite3.Error, OSError, ValueError) as exc:
            self.carita.setText("⚠️")
            self.cantidad.setText("Datos no disponibles")
            self.estado.setText("No se pudo leer el historial")
            self.barra.setValue(0)
            self.semana.setText(str(exc))

    def ver_semana(self):
        ventana = VentanaMetricas(self.lector, self)
        ventana.exec()