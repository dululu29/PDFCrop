#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
pdfcrop_gui.py

PyQt6 GUI for cropping PDF page margins while keeping the output as PDF.

Preview is rendered from the PDF page for display only.
Saving writes a new PDF with updated CropBox values.

Install:
    pip install PyQt6 pymupdf pillow

Run:
    python pdfcrop_gui.py
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, Optional, Tuple

try:
    import pymupdf as fitz
except ImportError:  # older alias still used in many environments
    import fitz

from PyQt6.QtCore import Qt, QRectF, pyqtSignal, QPointF
from PyQt6.QtGui import QPainter, QColor, QPen, QPixmap, QImage, QIntValidator
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSlider,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QCheckBox,
)


def pixmap_from_fitz(pix) -> QPixmap:
    """
    Convert a PyMuPDF pixmap to QPixmap.
    """
    if pix.alpha:
        fmt = QImage.Format.Format_RGBA8888
    else:
        fmt = QImage.Format.Format_RGB888

    qimg = QImage(
        pix.samples,
        pix.width,
        pix.height,
        pix.stride,
        fmt,
    ).copy()
    return QPixmap.fromImage(qimg)


class CropPreviewLabel(QLabel):
    marginsChanged = pyqtSignal(int, int, int, int)

    HANDLE_NONE = 0
    HANDLE_LEFT = 1
    HANDLE_RIGHT = 2
    HANDLE_TOP = 3
    HANDLE_BOTTOM = 4

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMouseTracking(True)

        self.base_pixmap: Optional[QPixmap] = None
        self.display_rect = QRectF()

        self.page_w = 0.0  # in PDF points
        self.page_h = 0.0  # in PDF points

        self.left_margin = 0
        self.right_margin = 0
        self.top_margin = 0
        self.bottom_margin = 0

        self.active_handle = self.HANDLE_NONE
        self.handle_tol = 8.0

    def set_page(self, pixmap: QPixmap, page_w: float, page_h: float):
        self.base_pixmap = pixmap
        self.page_w = max(1.0, float(page_w))
        self.page_h = max(1.0, float(page_h))
        self.update()

    def set_margins(self, left: int, right: int, top: int, bottom: int):
        self.left_margin = left
        self.right_margin = right
        self.top_margin = top
        self.bottom_margin = bottom
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.base_pixmap:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        widget_w = self.width()
        widget_h = self.height()
        pix_w = self.base_pixmap.width()
        pix_h = self.base_pixmap.height()

        scale = min(widget_w / pix_w, widget_h / pix_h)
        draw_w = pix_w * scale
        draw_h = pix_h * scale
        x0 = (widget_w - draw_w) / 2
        y0 = (widget_h - draw_h) / 2
        self.display_rect = QRectF(x0, y0, draw_w, draw_h)

        painter.drawPixmap(int(x0), int(y0), int(draw_w), int(draw_h), self.base_pixmap)

        sx = draw_w / self.page_w
        sy = draw_h / self.page_h

        crop_left_x = x0 + self.left_margin * sx
        crop_right_x = x0 + draw_w - self.right_margin * sx
        crop_top_y = y0 + self.top_margin * sy
        crop_bottom_y = y0 + draw_h - self.bottom_margin * sy

        shade = QColor(0, 0, 0, 75)
        painter.fillRect(QRectF(x0, y0, crop_left_x - x0, draw_h), shade)
        painter.fillRect(QRectF(crop_right_x, y0, x0 + draw_w - crop_right_x, draw_h), shade)
        painter.fillRect(QRectF(crop_left_x, y0, crop_right_x - crop_left_x, crop_top_y - y0), shade)
        painter.fillRect(QRectF(crop_left_x, crop_bottom_y, crop_right_x - crop_left_x, y0 + draw_h - crop_bottom_y), shade)

        pen = QPen(QColor(255, 50, 50), 2)
        painter.setPen(pen)
        painter.drawRect(QRectF(crop_left_x, crop_top_y, crop_right_x - crop_left_x, crop_bottom_y - crop_top_y))

    def _handle_at_pos(self, pos) -> int:
        if not self.base_pixmap:
            return self.HANDLE_NONE

        x0 = self.display_rect.left()
        y0 = self.display_rect.top()
        draw_w = self.display_rect.width()
        draw_h = self.display_rect.height()

        sx = draw_w / self.page_w
        sy = draw_h / self.page_h

        crop_left_x = x0 + self.left_margin * sx
        crop_right_x = x0 + draw_w - self.right_margin * sx
        crop_top_y = y0 + self.top_margin * sy
        crop_bottom_y = y0 + draw_h - self.bottom_margin * sy

        x = pos.x()
        y = pos.y()

        if crop_top_y <= y <= crop_bottom_y:
            if abs(x - crop_left_x) <= self.handle_tol:
                return self.HANDLE_LEFT
            if abs(x - crop_right_x) <= self.handle_tol:
                return self.HANDLE_RIGHT

        if crop_left_x <= x <= crop_right_x:
            if abs(y - crop_top_y) <= self.handle_tol:
                return self.HANDLE_TOP
            if abs(y - crop_bottom_y) <= self.handle_tol:
                return self.HANDLE_BOTTOM

        return self.HANDLE_NONE

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.active_handle = self._handle_at_pos(event.position().toPoint())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        pos = event.position().toPoint()
        handle = self._handle_at_pos(pos)

        if handle in (self.HANDLE_LEFT, self.HANDLE_RIGHT):
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        elif handle in (self.HANDLE_TOP, self.HANDLE_BOTTOM):
            self.setCursor(Qt.CursorShape.SizeVerCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)

        if self.active_handle == self.HANDLE_NONE or not (event.buttons() & Qt.MouseButton.LeftButton):
            super().mouseMoveEvent(event)
            return

        x0 = self.display_rect.left()
        y0 = self.display_rect.top()
        draw_w = self.display_rect.width()
        draw_h = self.display_rect.height()

        sx = draw_w / self.page_w
        sy = draw_h / self.page_h

        left = self.left_margin
        right = self.right_margin
        top = self.top_margin
        bottom = self.bottom_margin

        x = pos.x()
        y = pos.y()

        if self.active_handle == self.HANDLE_LEFT:
            left = int(round((x - x0) / sx))
            left = max(0, min(left, int(self.page_w) - right - 1))
        elif self.active_handle == self.HANDLE_RIGHT:
            right = int(round((x0 + draw_w - x) / sx))
            right = max(0, min(right, int(self.page_w) - left - 1))
        elif self.active_handle == self.HANDLE_TOP:
            top = int(round((y - y0) / sy))
            top = max(0, min(top, int(self.page_h) - bottom - 1))
        elif self.active_handle == self.HANDLE_BOTTOM:
            bottom = int(round((y0 + draw_h - y) / sy))
            bottom = max(0, min(bottom, int(self.page_h) - top - 1))

        self.marginsChanged.emit(left, right, top, bottom)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.active_handle = self.HANDLE_NONE
        super().mouseReleaseEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Crop Tool (PyQt6)")
        self.resize(1500, 980)

        self.pdf_path: Optional[Path] = None
        self.doc = None
        self.base_cropboxes = []
        self.page_margins: Dict[int, Tuple[int, int, int, int]] = {}
        self.current_page_index = 0
        self.render_zoom = 2.0  # preview only

        self._syncing = False
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)

        # Left: big preview
        self.preview_label = CropPreviewLabel()
        self.preview_label.setMinimumSize(850, 760)
        self.preview_label.setStyleSheet("background: #222;")
        self.preview_label.marginsChanged.connect(self.on_preview_dragged)

        left_wrap = QScrollArea()
        left_wrap.setWidgetResizable(True)
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.addWidget(self.preview_label)
        left_wrap.setWidget(left_container)

        # Right panel
        right_panel = QVBoxLayout()

        btn_row = QHBoxLayout()
        self.btn_open = QPushButton("Open PDF")
        self.btn_prev = QPushButton("Prev Page")
        self.btn_next = QPushButton("Next Page")
        btn_row.addWidget(self.btn_open)
        btn_row.addWidget(self.btn_prev)
        btn_row.addWidget(self.btn_next)

        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)

        page_row = QHBoxLayout()
        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, 1)
        self.page_spin.valueChanged.connect(self.on_page_spin_changed)
        self.btn_apply_prev = QPushButton("Apply Previous Page Crop")
        self.btn_reset_page = QPushButton("Reset This Page")
        self.btn_reset_all = QPushButton("Reset All Pages")
        self.btn_apply_prev.clicked.connect(self.apply_previous_page_crop)
        self.btn_reset_page.clicked.connect(self.reset_current_page_margins)
        self.btn_reset_all.clicked.connect(self.reset_all_margins)
        page_row.addWidget(QLabel("Page"))
        page_row.addWidget(self.page_spin)
        page_row.addWidget(self.btn_apply_prev)
        page_row.addWidget(self.btn_reset_page)
        page_row.addWidget(self.btn_reset_all)

        self.chk_apply_same_to_all = QCheckBox("Save 時把目前頁 margins 套用到所有頁")
        self.chk_apply_same_to_all.setChecked(False)

        control_box = QGroupBox("Crop Margins (PDF points)")
        control_layout = QGridLayout(control_box)

        (
            self.left_slider,
            self.left_spin,
            self.left_edit,
        ) = self._make_slider_spin_edit_triplet()

        (
            self.right_slider,
            self.right_spin,
            self.right_edit,
        ) = self._make_slider_spin_edit_triplet()

        (
            self.top_slider,
            self.top_spin,
            self.top_edit,
        ) = self._make_slider_spin_edit_triplet()

        (
            self.bottom_slider,
            self.bottom_spin,
            self.bottom_edit,
        ) = self._make_slider_spin_edit_triplet()

        control_layout.addWidget(QLabel("Left"), 0, 0)
        control_layout.addWidget(self.left_slider, 0, 1)
        control_layout.addWidget(self.left_spin, 0, 2)
        control_layout.addWidget(self.left_edit, 0, 3)

        control_layout.addWidget(QLabel("Right"), 1, 0)
        control_layout.addWidget(self.right_slider, 1, 1)
        control_layout.addWidget(self.right_spin, 1, 2)
        control_layout.addWidget(self.right_edit, 1, 3)

        control_layout.addWidget(QLabel("Top"), 2, 0)
        control_layout.addWidget(self.top_slider, 2, 1)
        control_layout.addWidget(self.top_spin, 2, 2)
        control_layout.addWidget(self.top_edit, 2, 3)

        control_layout.addWidget(QLabel("Bottom"), 3, 0)
        control_layout.addWidget(self.bottom_slider, 3, 1)
        control_layout.addWidget(self.bottom_spin, 3, 2)
        control_layout.addWidget(self.bottom_edit, 3, 3)

        self.left_slider.valueChanged.connect(lambda v: self._triplet_changed("left", v))
        self.right_slider.valueChanged.connect(lambda v: self._triplet_changed("right", v))
        self.top_slider.valueChanged.connect(lambda v: self._triplet_changed("top", v))
        self.bottom_slider.valueChanged.connect(lambda v: self._triplet_changed("bottom", v))

        self.left_spin.valueChanged.connect(lambda v: self._triplet_changed("left", v))
        self.right_spin.valueChanged.connect(lambda v: self._triplet_changed("right", v))
        self.top_spin.valueChanged.connect(lambda v: self._triplet_changed("top", v))
        self.bottom_spin.valueChanged.connect(lambda v: self._triplet_changed("bottom", v))

        self.left_edit.editingFinished.connect(lambda: self._line_edit_finished("left"))
        self.right_edit.editingFinished.connect(lambda: self._line_edit_finished("right"))
        self.top_edit.editingFinished.connect(lambda: self._line_edit_finished("top"))
        self.bottom_edit.editingFinished.connect(lambda: self._line_edit_finished("bottom"))

        self.info_text = QLabel("Open a PDF first.")
        self.info_text.setWordWrap(True)

        self.result_label = QLabel()
        self.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.result_label.setMinimumHeight(260)
        self.result_label.setStyleSheet("background: #f0f0f0; border: 1px solid #ccc;")

        result_box = QGroupBox("Cropped Preview")
        result_layout = QVBoxLayout(result_box)
        result_layout.addWidget(self.info_text)
        result_layout.addWidget(self.result_label)

        save_row = QHBoxLayout()
        self.btn_save_current = QPushButton("Save Current Page PDF")
        self.btn_save_all = QPushButton("Save Cropped PDF")
        self.btn_save_current.clicked.connect(self.save_current_page_pdf)
        self.btn_save_all.clicked.connect(self.save_all_pages_pdf)
        save_row.addWidget(self.btn_save_current)
        save_row.addWidget(self.btn_save_all)

        hint = QLabel(
            "Tip: preview is raster only for display, but saving updates the PDF CropBox.\n"
            "You can drag the red box, use slider/spinbox, or type values directly in the rightmost fields."
        )
        hint.setWordWrap(True)

        right_panel.addLayout(btn_row)
        right_panel.addLayout(page_row)
        right_panel.addWidget(self.chk_apply_same_to_all)
        right_panel.addWidget(control_box)
        right_panel.addLayout(save_row)
        right_panel.addWidget(result_box)
        right_panel.addWidget(hint)
        right_panel.addStretch(1)

        root.addWidget(left_wrap, 3)
        root.addLayout(right_panel, 2)

        self._set_buttons_enabled(False)

    def _make_slider_spin_edit_triplet(self):
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(0, 0)

        spin = QSpinBox()
        spin.setRange(0, 0)
        spin.setMinimumWidth(80)

        edit = QLineEdit()
        edit.setPlaceholderText("type")
        edit.setMinimumWidth(72)
        edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        edit.setValidator(QIntValidator(0, 0, self))

        return slider, spin, edit

    def _set_triplet_value(self, slider: QSlider, spin: QSpinBox, edit: QLineEdit, value: int):
        slider.blockSignals(True)
        spin.blockSignals(True)
        edit.blockSignals(True)
        slider.setValue(value)
        spin.setValue(value)
        edit.setText(str(value))
        slider.blockSignals(False)
        spin.blockSignals(False)
        edit.blockSignals(False)

    def _set_buttons_enabled(self, enabled: bool):
        for w in (
            self.btn_prev,
            self.btn_next,
            self.page_spin,
            self.btn_apply_prev,
            self.btn_reset_page,
            self.btn_reset_all,
            self.btn_save_current,
            self.btn_save_all,
            self.left_slider,
            self.left_spin,
            self.left_edit,
            self.right_slider,
            self.right_spin,
            self.right_edit,
            self.top_slider,
            self.top_spin,
            self.top_edit,
            self.bottom_slider,
            self.bottom_spin,
            self.bottom_edit,
        ):
            w.setEnabled(enabled)

    def _update_nav_and_page_buttons(self):
        has_doc = self.doc is not None
        if not has_doc:
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.btn_apply_prev.setEnabled(False)
            return

        self.btn_prev.setEnabled(self.current_page_index > 0)
        self.btn_next.setEnabled(self.current_page_index < self.doc.page_count - 1)
        self.btn_apply_prev.setEnabled(self.current_page_index > 0)

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open PDF",
            "",
            "PDF Files (*.pdf)"
        )
        if not path:
            return

        try:
            self.close_current_doc()

            self.pdf_path = Path(path)
            self.doc = fitz.open(str(self.pdf_path))
            if self.doc.page_count < 1:
                raise ValueError("PDF has no pages.")

            self.base_cropboxes = [fitz.Rect(self.doc[i].cropbox) for i in range(self.doc.page_count)]
            self.page_margins = {i: (0, 0, 0, 0) for i in range(self.doc.page_count)}
            self.current_page_index = 0

            self.page_spin.blockSignals(True)
            self.page_spin.setRange(1, self.doc.page_count)
            self.page_spin.setValue(1)
            self.page_spin.blockSignals(False)

            self._set_buttons_enabled(True)
            self.load_current_page()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open PDF:\n{e}")
            self.close_current_doc()

    def close_current_doc(self):
        if self.doc is not None:
            self.doc.close()
        self.doc = None
        self.pdf_path = None
        self.base_cropboxes = []
        self.page_margins = {}
        self.current_page_index = 0
        self._set_buttons_enabled(False)

    def current_base_rect(self):
        return self.base_cropboxes[self.current_page_index]

    def current_margins(self):
        return self.page_margins.get(self.current_page_index, (0, 0, 0, 0))

    def load_current_page(self):
        if self.doc is None:
            return

        page = self.doc[self.current_page_index]
        base = self.current_base_rect()
        rotation = page.rotation

        if rotation != 0:
            QMessageBox.warning(
                self,
                "Rotation warning",
                f"Page {self.current_page_index + 1} has rotation={rotation}.\n"
                "This GUI is designed primarily for unrotated pages.\n"
                "For PowerPoint-exported figure PDFs this is usually 0."
            )

        matrix = fitz.Matrix(self.render_zoom, self.render_zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        qpix = pixmap_from_fitz(pix)

        self.preview_label.set_page(qpix, base.width, base.height)
        self._reset_control_ranges(base)
        self._load_page_margins_into_controls()
        self._update_nav_and_page_buttons()
        self.update_preview()

    def _reset_control_ranges(self, base_rect):
        max_w = max(0, int(round(base_rect.width)) - 1)
        max_h = max(0, int(round(base_rect.height)) - 1)

        for slider, spin, edit in (
            (self.left_slider, self.left_spin, self.left_edit),
            (self.right_slider, self.right_spin, self.right_edit),
        ):
            slider.setRange(0, max_w)
            spin.setRange(0, max_w)
            edit.setValidator(QIntValidator(0, max_w, self))

        for slider, spin, edit in (
            (self.top_slider, self.top_spin, self.top_edit),
            (self.bottom_slider, self.bottom_spin, self.bottom_edit),
        ):
            slider.setRange(0, max_h)
            spin.setRange(0, max_h)
            edit.setValidator(QIntValidator(0, max_h, self))

    def _load_page_margins_into_controls(self):
        left, right, top, bottom = self.current_margins()
        self._syncing = True
        self._set_triplet_value(self.left_slider, self.left_spin, self.left_edit, left)
        self._set_triplet_value(self.right_slider, self.right_spin, self.right_edit, right)
        self._set_triplet_value(self.top_slider, self.top_spin, self.top_edit, top)
        self._set_triplet_value(self.bottom_slider, self.bottom_spin, self.bottom_edit, bottom)
        self.preview_label.set_margins(left, right, top, bottom)
        self._syncing = False

    def save_current_control_values(self):
        if self.doc is None:
            return
        self.page_margins[self.current_page_index] = (
            self.left_spin.value(),
            self.right_spin.value(),
            self.top_spin.value(),
            self.bottom_spin.value(),
        )

    def _line_edit_finished(self, which: str):
        if self._syncing or self.doc is None:
            return

        edit_map = {
            "left": self.left_edit,
            "right": self.right_edit,
            "top": self.top_edit,
            "bottom": self.bottom_edit,
        }
        spin_map = {
            "left": self.left_spin,
            "right": self.right_spin,
            "top": self.top_spin,
            "bottom": self.bottom_spin,
        }

        edit = edit_map[which]
        fallback = spin_map[which].value()
        text = edit.text().strip()

        if text == "":
            value = fallback
        else:
            try:
                value = int(text)
            except ValueError:
                value = fallback

        self._triplet_changed(which, value)

    def _triplet_changed(self, which: str, value: int):
        if self._syncing or self.doc is None:
            return

        self._syncing = True

        if which == "left":
            self._set_triplet_value(self.left_slider, self.left_spin, self.left_edit, value)
        elif which == "right":
            self._set_triplet_value(self.right_slider, self.right_spin, self.right_edit, value)
        elif which == "top":
            self._set_triplet_value(self.top_slider, self.top_spin, self.top_edit, value)
        elif which == "bottom":
            self._set_triplet_value(self.bottom_slider, self.bottom_spin, self.bottom_edit, value)

        self._enforce_valid_ranges()
        self.save_current_control_values()
        self.update_preview()
        self._syncing = False

    def _enforce_valid_ranges(self):
        base = self.current_base_rect()
        w = max(1, int(round(base.width)))
        h = max(1, int(round(base.height)))

        left = self.left_spin.value()
        right = self.right_spin.value()
        top = self.top_spin.value()
        bottom = self.bottom_spin.value()

        if left + right >= w:
            if self.sender() in (self.left_slider, self.left_spin, self.left_edit):
                right = max(0, w - left - 1)
            else:
                left = max(0, w - right - 1)

        if top + bottom >= h:
            if self.sender() in (self.top_slider, self.top_spin, self.top_edit):
                bottom = max(0, h - top - 1)
            else:
                top = max(0, h - bottom - 1)

        self._set_triplet_value(self.left_slider, self.left_spin, self.left_edit, left)
        self._set_triplet_value(self.right_slider, self.right_spin, self.right_edit, right)
        self._set_triplet_value(self.top_slider, self.top_spin, self.top_edit, top)
        self._set_triplet_value(self.bottom_slider, self.bottom_spin, self.bottom_edit, bottom)

        self.preview_label.set_margins(left, right, top, bottom)

    def on_preview_dragged(self, left: int, right: int, top: int, bottom: int):
        if self._syncing:
            return
        self._syncing = True
        self._set_triplet_value(self.left_slider, self.left_spin, self.left_edit, left)
        self._set_triplet_value(self.right_slider, self.right_spin, self.right_edit, right)
        self._set_triplet_value(self.top_slider, self.top_spin, self.top_edit, top)
        self._set_triplet_value(self.bottom_slider, self.bottom_spin, self.bottom_edit, bottom)
        self.preview_label.set_margins(left, right, top, bottom)
        self.save_current_control_values()
        self.update_preview()
        self._syncing = False

    def on_page_spin_changed(self, value: int):
        if self.doc is None:
            return
        self.save_current_control_values()
        self.current_page_index = value - 1
        self.load_current_page()

    def prev_page(self):
        if self.doc is None or self.current_page_index <= 0:
            return
        self.save_current_control_values()
        self.current_page_index -= 1
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(self.current_page_index + 1)
        self.page_spin.blockSignals(False)
        self.load_current_page()

    def next_page(self):
        if self.doc is None or self.current_page_index >= self.doc.page_count - 1:
            return
        self.save_current_control_values()
        self.current_page_index += 1
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(self.current_page_index + 1)
        self.page_spin.blockSignals(False)
        self.load_current_page()

    def apply_previous_page_crop(self):
        if self.doc is None or self.current_page_index <= 0:
            return

        prev_margins = self.page_margins.get(self.current_page_index - 1, (0, 0, 0, 0))
        self.page_margins[self.current_page_index] = prev_margins
        self._load_page_margins_into_controls()
        self.update_preview()

    def reset_current_page_margins(self):
        if self.doc is None:
            return
        self.page_margins[self.current_page_index] = (0, 0, 0, 0)
        self._load_page_margins_into_controls()
        self.update_preview()

    def reset_all_margins(self):
        if self.doc is None:
            return
        self.page_margins = {i: (0, 0, 0, 0) for i in range(self.doc.page_count)}
        self._load_page_margins_into_controls()
        self.update_preview()

    def get_current_crop_rect(self) -> fitz.Rect:
        base = self.current_base_rect()
        left, right, top, bottom = self.current_margins()

        new_rect = fitz.Rect(
            base.x0 + left,
            base.y0 + top,
            base.x1 - right,
            base.y1 - bottom,
        )
        if new_rect.is_empty or new_rect.width <= 0 or new_rect.height <= 0:
            raise ValueError("Crop margins are too large.")
        return new_rect

    def update_preview(self):
        if self.doc is None:
            return

        try:
            base = self.current_base_rect()
            crop = self.get_current_crop_rect()

            scale_x = 500 / max(1.0, base.width)
            scale_y = 260 / max(1.0, base.height)
            scale = min(scale_x, scale_y, 1.0)

            preview_w = max(1, int(round(crop.width * scale)))
            preview_h = max(1, int(round(crop.height * scale)))
            preview = QPixmap(preview_w, preview_h)
            preview.fill(QColor("white"))

            painter = QPainter(preview)
            painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

            if self.preview_label.base_pixmap is not None:
                full_pm = self.preview_label.base_pixmap
                full_w = full_pm.width()
                full_h = full_pm.height()

                # crop position relative to the base cropbox rectangle
                x_ratio = (crop.x0 - base.x0) / max(1.0, base.width)
                y_ratio = (crop.y0 - base.y0) / max(1.0, base.height)
                w_ratio = crop.width / max(1.0, base.width)
                h_ratio = crop.height / max(1.0, base.height)

                src_x = int(round(full_w * x_ratio))
                src_y = int(round(full_h * y_ratio))
                src_w = int(round(full_w * w_ratio))
                src_h = int(round(full_h * h_ratio))

                painter.drawPixmap(
                    QRectF(0, 0, preview_w, preview_h),
                    full_pm,
                    QRectF(src_x, src_y, src_w, src_h),
                )
            painter.end()

            self.result_label.setPixmap(preview)

            left, right, top, bottom = self.current_margins()
            self.preview_label.set_margins(left, right, top, bottom)

            self.info_text.setText(
                f"Page {self.current_page_index + 1} / {self.doc.page_count}\n"
                f"Base CropBox: {base.width:.1f} × {base.height:.1f} pt\n"
                f"New CropBox:  {crop.width:.1f} × {crop.height:.1f} pt\n"
                f"Margins: L={left}  R={right}  T={top}  B={bottom}"
            )
        except Exception as e:
            self.result_label.clear()
            self.info_text.setText(f"Preview failed: {e}")

    def _apply_crop_to_doc(self, out_doc, apply_same_to_all: bool):
        for i in range(out_doc.page_count):
            page = out_doc[i]
            base = fitz.Rect(self.base_cropboxes[i])

            if apply_same_to_all:
                margins = self.current_margins()
            else:
                margins = self.page_margins.get(i, (0, 0, 0, 0))

            left, right, top, bottom = margins
            new_rect = fitz.Rect(
                base.x0 + left,
                base.y0 + top,
                base.x1 - right,
                base.y1 - bottom,
            )

            if new_rect.is_empty or new_rect.width <= 0 or new_rect.height <= 0:
                raise ValueError(f"Page {i + 1}: crop margins are too large.")

            page.set_cropbox(new_rect)

    def save_all_pages_pdf(self):
        if self.doc is None or self.pdf_path is None:
            return

        self.save_current_control_values()

        default_path = str(self.pdf_path.with_name(self.pdf_path.stem + "_cropped.pdf"))
        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Cropped PDF",
            default_path,
            "PDF Files (*.pdf)"
        )
        if not out_path:
            return

        try:
            out_doc = fitz.open(str(self.pdf_path))
            self._apply_crop_to_doc(out_doc, apply_same_to_all=self.chk_apply_same_to_all.isChecked())
            out_doc.save(out_path)
            out_doc.close()
            QMessageBox.information(self, "Saved", f"Saved cropped PDF:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Save failed:\n{e}")

    def save_current_page_pdf(self):
        if self.doc is None or self.pdf_path is None:
            return

        self.save_current_control_values()

        default_name = f"{self.pdf_path.stem}_page{self.current_page_index + 1}_cropped.pdf"
        default_path = str(self.pdf_path.with_name(default_name))

        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Current Page PDF",
            default_path,
            "PDF Files (*.pdf)"
        )
        if not out_path:
            return

        try:
            src_doc = fitz.open(str(self.pdf_path))
            out_doc = fitz.open()
            out_doc.insert_pdf(src_doc, from_page=self.current_page_index, to_page=self.current_page_index)
            src_doc.close()

            base = fitz.Rect(self.base_cropboxes[self.current_page_index])
            left, right, top, bottom = self.current_margins()
            new_rect = fitz.Rect(
                base.x0 + left,
                base.y0 + top,
                base.x1 - right,
                base.y1 - bottom,
            )

            if new_rect.is_empty or new_rect.width <= 0 or new_rect.height <= 0:
                raise ValueError("Crop margins are too large.")

            out_doc[0].set_cropbox(new_rect)
            out_doc.save(out_path)
            out_doc.close()
            QMessageBox.information(self, "Saved", f"Saved current page PDF:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Save failed:\n{e}")

    def closeEvent(self, event):
        self.close_current_doc()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
