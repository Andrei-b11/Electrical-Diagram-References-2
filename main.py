import sys
import re
import os
import json
import subprocess
import ctypes
import ctypes.wintypes
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QTableWidget, QTableWidgetItem, QTextEdit, QSplitter,
                             QHeaderView, QMessageBox, QProgressDialog, QProgressBar, QComboBox,
                             QLineEdit, QGroupBox, QFormLayout, QSpinBox, QDialog,
                             QGraphicsView, QGraphicsScene, QGraphicsLineItem,
                             QGraphicsRectItem, QSlider, QScrollArea, QFrame,
                             QRadioButton, QButtonGroup, QTabWidget, QListWidget,
                             QListWidgetItem, QAbstractItemView, QCheckBox, QMenu,
                             QSizePolicy, QStackedWidget, QGridLayout)
from PyQt5.QtCore import Qt, QRectF, QPointF, QLineF, QSize, QPropertyAnimation, QEasingCurve, QThread, pyqtSignal
from PyQt5.QtGui import (QFont, QPixmap, QImage, QPen, QColor, QBrush, QPainter, 
                         QIcon, QLinearGradient, QPalette, QFontDatabase, QPainterPath, QRegion, QTransform)
from PyQt5.QtSvg import QSvgRenderer
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import DictionaryObject, NameObject, ArrayObject, NumberObject, create_string_object


def load_svg_icons():
    """Cargar iconos SVG desde icons.json"""
    try:
        icons_path = os.path.join(get_app_path(), 'icons.json')
        if os.path.exists(icons_path):
            with open(icons_path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading icons: {e}")
    return {}

def create_svg_icon(svg_code, size=24):
    """Crear un QLabel con icono SVG"""
    label = QLabel()
    
    # Crear un QPixmap desde el código SVG
    svg_bytes = svg_code.encode('utf-8')
    pixmap = QPixmap()
    
    # Usar QPainter para renderizar el SVG
    from PyQt5.QtSvg import QSvgRenderer
    renderer = QSvgRenderer(svg_bytes)
    
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    
    label.setPixmap(pixmap)
    label.setAlignment(Qt.AlignCenter)
    label.setStyleSheet("background: transparent; border: none;")
    
    return label

def get_app_path():
    """Obtener la ruta del directorio de la aplicación"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Cargar iconos SVG al inicio
SVG_ICONS = load_svg_icons()

# ─────────────────────────────────────────────
#  WORKER THREADS
# ─────────────────────────────────────────────

class DetectionWorker(QThread):
    """Worker thread para detección de referencias"""
    progress_updated = pyqtSignal(int, str)  # valor, texto
    finished_signal = pyqtSignal(list, dict)  # referencias, all_references
    error_signal = pyqtSignal(str)
    
    def __init__(self, pdf_paths, pattern, groups_order):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.pattern = pattern
        self.groups_order = groups_order
        self.canceled = False
        
    def cancel(self):
        self.canceled = True
        
    def run(self):
        try:
            import fitz
            import re
            
            # Calcular total de páginas
            total_pages = 0
            pdf_page_counts = {}
            for i, pp in enumerate(self.pdf_paths):
                if self.canceled: return
                self.progress_updated.emit(int((i / len(self.pdf_paths)) * 20), 
                                         f'Calculating pages... ({i+1}/{len(self.pdf_paths)})')
                try:
                    td = fitz.open(pp)
                    pdf_page_counts[pp] = len(td)
                    total_pages += len(td)
                    td.close()
                except:
                    pdf_page_counts[pp] = 0
            
            if self.canceled: return
            
            # Análisis de referencias
            references = []
            all_references = {}
            current_page = 0
            
            for pdf_path in self.pdf_paths:
                if self.canceled: return
                pdf_name = os.path.basename(pdf_path)
                
                try:
                    doc = fitz.open(pdf_path)
                    pdf_refs = []
                    
                    for page_num in range(len(doc)):
                        if self.canceled: return
                        
                        current_page += 1
                        progress_pct = 20 + int((current_page / total_pages) * 80)
                        self.progress_updated.emit(progress_pct, 
                                                 f'Analyzing {pdf_name} - Page {page_num+1}')
                        
                        page = doc[page_num]
                        text = page.get_text()
                        
                        # Buscar referencias usando el patrón
                        matches = list(re.finditer(self.pattern, text, re.IGNORECASE | re.MULTILINE))
                        ref_pos_used = {}
                        
                        for match in matches:
                            if self.canceled: return
                            
                            full = match.group(0)
                            g = [match.group(i) if match.lastindex and match.lastindex >= i else '' for i in range(1,4)]
                            
                            # Procesar grupos según el orden
                            pr, cr, rr = '', '', ''
                            for i, gn in enumerate(self.groups_order):
                                v = g[i] if i < len(g) else ''
                                if gn == 'página': pr = v
                                elif gn == 'columna': cr = v
                                elif gn == 'fila': rr = v
                            
                            # Contexto
                            s = max(0, match.start()-30)
                            e = min(len(text), match.end()+30)
                            ctx = text[s:e].replace('\n',' ').strip()
                            
                            # Buscar coordenadas
                            instances = page.search_for(full)
                            if instances:
                                idx = ref_pos_used.get(full, 0)
                                rect = instances[idx] if idx < len(instances) else instances[-1]
                                ref_pos_used[full] = idx + 1
                                
                                ref_data = {
                                    'full': full,
                                    'page': pr,
                                    'column': cr,
                                    'row': rr,
                                    'context': ctx,
                                    'pdf_page': page_num,
                                    'coordinates': [rect.x0, rect.y0, rect.x1, rect.y1],
                                    'instance': idx + 1,
                                    'pdf_path': pdf_path,
                                    'pdf_name': pdf_name
                                }
                                references.append(ref_data)
                                pdf_refs.append(ref_data)
                    
                    doc.close()
                    all_references[pdf_path] = pdf_refs
                    
                except Exception as e:
                    print(f"Error processing {pdf_path}: {e}")
                    continue
            
            if not self.canceled:
                self.finished_signal.emit(references, all_references)
                
        except Exception as e:
            self.error_signal.emit(str(e))


class PDFGenerationWorker(QThread):
    """Worker thread para generación de PDFs"""
    progress_updated = pyqtSignal(int, str)  # valor, texto
    finished_signal = pyqtSignal(list, int)  # archivos generados, total_links
    error_signal = pyqtSignal(str)
    
    def __init__(self, detector_instance, all_references, keep_original_name, output_dir):
        super().__init__()
        self.detector = detector_instance
        self.all_references = all_references
        self.keep_original_name = keep_original_name
        self.output_dir = output_dir
        self.canceled = False
        
    def cancel(self):
        self.canceled = True
        
    def run(self):
        try:
            generated = []
            total_links = 0
            
            for i, (pdf_path, references) in enumerate(self.all_references.items()):
                if self.canceled: return
                
                pdf_name = os.path.basename(pdf_path)
                progress_pct = int((i / len(self.all_references)) * 100)
                self.progress_updated.emit(progress_pct, f'Generating {pdf_name}...')
                
                # Generar PDF individual usando el método del detector
                output_path = self.detector._generate_single_pdf(pdf_path, references, self.keep_original_name, self.output_dir)
                if output_path:
                    generated.append(output_path)
                    total_links += len(references)
            
            if not self.canceled:
                self.finished_signal.emit(generated, total_links)
                
        except Exception as e:
            self.error_signal.emit(str(e))


# ─────────────────────────────────────────────
#  DESIGN TOKENS
# ─────────────────────────────────────────────
COLORS = {
    'bg_base':        '#FFFFFF',  # fondo - blanco
    'bg_surface':     '#F3F4F6',  # panel - gris claro
    'bg_elevated':    '#F5F5F5',  # sidebar - gris claro
    'bg_hover':       '#F3F4F6',  # hover gris claro
    'border':         '#FF6B2B',  # borde naranja para la ventana
    'border_focus':   '#FF6B2B',  # accent como foco
    'text_primary':   '#1F2937',  # texto oscuro para fondo blanco
    'text_secondary': '#6B7280',
    'text_muted':     '#9CA3AF',
    'accent':         '#FF6B2B',  # accent - color secundario predominante
    'accent_hover':   '#FF8A4D',  # accent más claro para hover
    'accent_dim':     '#4D1F0B',  # accent oscuro
    'success':        '#10B981',
    'success_dim':    '#064E3B',
    'warning':        '#F59E0B',
    'warning_dim':    '#451A03',
    'danger':         '#EF4444',
    'purple':         '#8B5CF6',
    'purple_dim':     '#3B0764',
}

APP_STYLESHEET = f"""
/* ── Global ────────────────────────────────── */
* {{
    font-family: 'Segoe UI', 'SF Pro Display', sans-serif;
    outline: none;
}}

QMainWindow, QDialog {{
    background-color: {COLORS['bg_base']};
}}

QWidget {{
    background-color: transparent;
    color: {COLORS['text_primary']};
    font-size: 13px;
}}

/* ── Scrollbars ─────────────────────────────── */
QScrollBar:vertical {{
    background: {COLORS['bg_surface']};
    width: 6px;
    border-radius: 3px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['border']};
    border-radius: 3px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLORS['text_muted']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background: {COLORS['bg_surface']};
    height: 6px;
    border-radius: 3px;
}}
QScrollBar::handle:horizontal {{
    background: {COLORS['border']};
    border-radius: 3px;
    min-width: 24px;
}}

/* ── Inputs ─────────────────────────────────── */
QLineEdit {{
    background: {COLORS['bg_elevated']};
    border: 1.5px solid {COLORS['border']};
    border-radius: 20px;
    padding: 8px 16px;
    color: {COLORS['text_primary']};
    font-size: 13px;
    selection-background-color: {COLORS['accent']};
}}
QLineEdit:focus {{
    border-color: {COLORS['accent']};
    background: {COLORS['bg_hover']};
}}
QLineEdit:disabled {{
    color: {COLORS['text_muted']};
    border-color: {COLORS['bg_elevated']};
}}
QLineEdit::placeholder {{
    color: {COLORS['text_muted']};
}}

QSpinBox {{
    background: {COLORS['bg_elevated']};
    border: 1.5px solid {COLORS['border']};
    border-radius: 8px;
    padding: 7px 10px;
    color: {COLORS['text_primary']};
    font-size: 13px;
    min-width: 72px;
}}
QSpinBox:focus {{
    border-color: {COLORS['accent']};
}}
QSpinBox::up-button {{
    background: {COLORS['bg_hover']};
    border: none;
    border-left: 1px solid {COLORS['border']};
    border-top-right-radius: 7px;
    width: 22px;
}}
QSpinBox::down-button {{
    background: {COLORS['bg_hover']};
    border: none;
    border-left: 1px solid {COLORS['border']};
    border-bottom-right-radius: 7px;
    width: 22px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
    background: {COLORS['accent_dim']};
}}
QSpinBox::up-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-bottom: 5px solid {COLORS['text_secondary']};
    margin: 0 5px;
}}
QSpinBox::down-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid {COLORS['text_secondary']};
    margin: 0 5px;
}}

QComboBox {{
    background: {COLORS['bg_elevated']};
    border: 1.5px solid {COLORS['border']};
    border-radius: 20px;
    padding: 0px 36px 0px 16px;  /* Eliminar padding vertical completamente */
    color: {COLORS['text_primary']};
    font-size: 12px;
    min-width: 120px;
    height: 28px;  /* Altura fija en lugar de min-height */
    line-height: 28px;  /* Line-height igual a la altura para centrar texto */
}}
QComboBox:focus, QComboBox:on {{
    border-color: {COLORS['accent']};
}}
QComboBox::drop-down {{
    border: none;
    width: 32px;
}}
QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid {COLORS['text_secondary']};
    margin-right: 10px;
}}
QComboBox QAbstractItemView {{
    background: {COLORS['bg_elevated']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['accent_dim']};
    selection-color: {COLORS['accent_hover']};
    padding: 4px;
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    padding: 8px 12px;
    border-radius: 4px;
    min-height: 28px;
}}

/* ── Checkboxes ──────────────────────────────── */
QCheckBox {{
    color: {COLORS['text_secondary']};
    font-size: 12px;
    spacing: 8px;
}}
QCheckBox:hover {{
    color: {COLORS['text_primary']};
}}
QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border: 1.5px solid {COLORS['border']};
    border-radius: 4px;
    background: {COLORS['bg_elevated']};
}}
QCheckBox::indicator:checked {{
    background: {COLORS['accent']};
    border-color: {COLORS['accent']};
}}
QCheckBox::indicator:hover {{
    border-color: {COLORS['accent']};
}}

/* ── Radio Buttons ───────────────────────────── */
QRadioButton {{
    color: {COLORS['text_secondary']};
    font-size: 12px;
    spacing: 8px;
}}
QRadioButton:hover {{
    color: {COLORS['text_primary']};
}}
QRadioButton::indicator {{
    width: 16px;
    height: 16px;
    border: 1.5px solid {COLORS['border']};
    border-radius: 8px;
    background: {COLORS['bg_elevated']};
}}
QRadioButton::indicator:checked {{
    background: {COLORS['accent']};
    border-color: {COLORS['accent']};
}}

/* ── Tables ─────────────────────────────────── */
QTableWidget {{
    background: {COLORS['bg_surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
    gridline-color: {COLORS['bg_elevated']};
    color: {COLORS['text_primary']};
    font-size: 12px;
    selection-background-color: {COLORS['accent_dim']};
    alternate-background-color: {COLORS['bg_base']};
}}
QTableWidget::item {{
    padding: 10px 12px;
    border: none;
}}
QTableWidget::item:selected {{
    background: {COLORS['accent_dim']};
    color: {COLORS['accent_hover']};
}}
QHeaderView::section {{
    background: {COLORS['bg_elevated']};
    color: {COLORS['text_secondary']};
    padding: 10px 12px;
    border: none;
    border-bottom: 1px solid {COLORS['border']};
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
QHeaderView {{
    background: transparent;
    border: none;
}}

/* ── TextEdit ────────────────────────────────── */
QTextEdit {{
    background: {COLORS['bg_surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
    color: {COLORS['text_primary']};
    padding: 12px;
    font-size: 13px;
    font-family: 'Consolas', 'JetBrains Mono', monospace;
}}

/* ── Status Bar ──────────────────────────────── */
QStatusBar {{
    background: {COLORS['bg_surface']};
    color: {COLORS['text_muted']};
    border-top: 1px solid {COLORS['border']};
    padding: 4px 12px;
    font-size: 12px;
}}
QStatusBar::item {{
    border: none;
}}

/* ── Splitter ────────────────────────────────── */
QSplitter::handle {{
    background: {COLORS['border']};
}}
QSplitter::handle:horizontal {{
    width: 1px;
}}
QSplitter::handle:vertical {{
    height: 1px;
}}
QSplitter::handle:hover {{
    background: {COLORS['accent']};
}}

/* ── Tooltips ────────────────────────────────── */
QToolTip {{
    background: {COLORS['bg_elevated']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}}

/* ── List Widget ─────────────────────────────── */
QListWidget {{
    background: {COLORS['bg_surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
    color: {COLORS['text_primary']};
    padding: 4px;
    font-size: 12px;
    outline: none;
}}
QListWidget::item {{
    padding: 2px;  /* Padding mínimo */
    border-radius: 8px;
    margin: 2px;  /* Margen entre items */
    background: transparent;  /* Fondo transparente para widgets personalizados */
    min-height: 40px;  /* Altura mínima garantizada */
}}
QListWidget::item:selected {{
    background: {COLORS['accent_dim']};
    color: {COLORS['accent_hover']};
}}
QListWidget::item:hover:!selected {{
    background: {COLORS['bg_elevated']};
}}

/* ── Slider ──────────────────────────────────── */
QSlider::groove:horizontal {{
    height: 4px;
    background: {COLORS['border']};
    border-radius: 2px;
}}
QSlider::handle:horizontal {{
    background: {COLORS['accent']};
    width: 14px;
    height: 14px;
    border-radius: 7px;
    margin: -5px 0;
}}
QSlider::sub-page:horizontal {{
    background: {COLORS['accent']};
    border-radius: 2px;
}}

/* ── Menu ────────────────────────────────────── */
QMenu {{
    background: {COLORS['bg_elevated']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
    padding: 6px;
}}
QMenu::item {{
    color: {COLORS['text_primary']};
    padding: 9px 18px;
    border-radius: 6px;
    font-size: 13px;
}}
QMenu::item:selected {{
    background: {COLORS['accent_dim']};
    color: {COLORS['accent_hover']};
}}
QMenu::separator {{
    height: 1px;
    background: {COLORS['border']};
    margin: 4px 8px;
}}
"""

def make_btn(text, color='accent', size='md', icon=''):
    """Factory for styled buttons."""
    btn = QPushButton(f"{icon}  {text}".strip() if icon else text)
    
    palette = {
        'accent':  (COLORS['accent'],       COLORS['accent_hover'],  '#1D4ED8',  '#FFFFFF'),
        'success': (COLORS['success'],       '#34D399',               '#065F46',  '#FFFFFF'),
        'warning': (COLORS['warning'],       '#FCD34D',               '#78350F',  '#0A0E1A'),
        'danger':  (COLORS['danger'],        '#F87171',               '#991B1B',  '#FFFFFF'),
        'purple':  (COLORS['purple'],        '#A78BFA',               '#4C1D95',  '#FFFFFF'),
        'ghost':   (COLORS['bg_elevated'],   COLORS['bg_hover'],      COLORS['bg_hover'], COLORS['text_primary']),
        'outline': ('transparent',           COLORS['bg_elevated'],   COLORS['bg_hover'], COLORS['text_primary']),
    }
    bg, hover, press, txt = palette.get(color, palette['accent'])
    
    pad = {'sm': '6px 14px', 'md': '9px 18px', 'lg': '11px 24px'}.get(size, '9px 18px')
    radius = {'sm': '7px', 'md': '8px', 'lg': '10px'}.get(size, '8px')
    font = {'sm': '12px', 'md': '13px', 'lg': '14px'}.get(size, '13px')

    border = f'1.5px solid {COLORS["border"]}' if color in ('ghost', 'outline') else 'none'

    btn.setStyleSheet(f"""
        QPushButton {{
            background: {bg};
            color: {txt};
            padding: {pad};
            border: {border};
            border-radius: {radius};
            font-size: {font};
            font-weight: 600;
            letter-spacing: 0.2px;
        }}
        QPushButton:hover {{
            background: {hover};
        }}
        QPushButton:pressed {{
            background: {press};
        }}
        QPushButton:disabled {{
            background: {COLORS['bg_elevated']};
            color: {COLORS['text_muted']};
            border: 1px solid {COLORS['bg_elevated']};
        }}
    """)
    return btn


def make_card(title='', accent_color=None):
    """Create a card frame with optional title."""
    frame = QFrame()
    frame.setObjectName('card')
    c = accent_color or COLORS['border']
    frame.setStyleSheet(f"""
        QFrame#card {{
            background: {COLORS['bg_surface']};
            border: 1px solid {COLORS['border']};
            border-top: 2px solid {c};
            border-radius: 12px;
            padding: 4px;
        }}
    """)
    return frame


def section_label(text):
    lbl = QLabel(text)
    lbl.setStyleSheet(f"""
        color: {COLORS['accent']};
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        padding: 0 2px;
    """)
    return lbl


def value_label(text='', color=None):
    lbl = QLabel(text)
    c = color or COLORS['text_primary']
    lbl.setStyleSheet(f"color: {c}; font-size: 13px;")
    return lbl


def divider():
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setStyleSheet(f"background: {COLORS['border']}; max-height: 1px; border: none;")
    return line


class Badge(QLabel):
    def __init__(self, text='', color='accent', parent=None):
        super().__init__(text, parent)
        colors = {
            'accent':  (COLORS['accent_dim'],   COLORS['accent']),
            'success': (COLORS['success_dim'],   COLORS['success']),
            'warning': (COLORS['warning_dim'],   COLORS['warning']),
            'muted':   (COLORS['bg_elevated'],   COLORS['text_muted']),
            'purple':  (COLORS['purple_dim'],    COLORS['purple']),
        }
        bg, fg = colors.get(color, colors['accent'])
        self.setStyleSheet(f"""
            background: {bg};
            color: {fg};
            border-radius: 10px;
            padding: 3px 10px;
            font-size: 11px;
            font-weight: 700;
        """)


# ─────────────────────────────────────────────
#  GRID EDITOR DIALOG
# ─────────────────────────────────────────────
class GridEditorDialog(QDialog):
    def __init__(self, parent=None, pdf_path=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.page_num = 0
        self.zoom_factor = 1.0
        self.column_lines = []
        self.row_lines = []
        self.current_mode = 'column'
        self.pdf_doc = None
        self.page_width = 0
        self.page_height = 0
        self.config_file = None
        self.init_ui()
        if pdf_path:
            self.load_pdf(pdf_path)
            self.load_saved_config()

    def init_ui(self):
        self.setWindowTitle('Grid Editor — Visual')
        self.setGeometry(80, 80, 1200, 820)
        self.setModal(True)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet(f"""
            QDialog {{
                background: {COLORS['bg_base']};
                border: 2px solid {COLORS['accent']};
                border-radius: 16px;
            }}
        """)
    def init_ui(self):
        self.setWindowTitle('Grid Editor — Visual')
        self.setGeometry(80, 80, 1200, 820)
        self.setModal(True)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet(f"""
            QDialog {{
                background: {COLORS['bg_base']};
                border: 2px solid {COLORS['accent']};
                border-radius: 16px;
            }}
        """)
        
        # Variables para arrastrar la ventana
        self.drag_position = None
        
        lay = QVBoxLayout(self)
        lay.setSpacing(0)
        lay.setContentsMargins(0, 0, 0, 0)

        # ── Barra de título personalizada ───────────────────────
        self.title_bar = QWidget()
        self.title_bar.setFixedHeight(58)
        self.title_bar.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_base']};
                border-bottom: 1px solid {COLORS['border']};
                border-top-left-radius: 14px;
                border-top-right-radius: 14px;
            }}
        """)
        
        title_layout = QHBoxLayout(self.title_bar)
        title_layout.setContentsMargins(20, 0, 16, 0)
        title_layout.setSpacing(12)
        
        # Título
        title_label = QLabel('Grid Editor — Visual')
        title_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 16px;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
        """)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        # Botones de control
        self.minimize_btn = QPushButton('−')
        self.minimize_btn.setFixedSize(32, 32)
        self.minimize_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 18px;
                font-weight: bold;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
            }}
        """)
        self.minimize_btn.clicked.connect(self.showMinimized)
        title_layout.addWidget(self.minimize_btn)
        
        self.close_btn = QPushButton('×')
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 18px;
                font-weight: bold;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background: #ef4444;
                color: white;
            }}
        """)
        self.close_btn.clicked.connect(self.reject)
        title_layout.addWidget(self.close_btn)
        
        # Instalar event filter para arrastrar
        self.title_bar.mousePressEvent = self.title_bar_mouse_press
        self.title_bar.mouseMoveEvent = self.title_bar_mouse_move
        
        lay.addWidget(self.title_bar)

        # ── Top toolbar ───────────────────────
        toolbar_widget = QWidget()
        toolbar_widget.setFixedHeight(70)
        toolbar_widget.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_surface']};
                border-bottom: 1px solid {COLORS['border']};
            }}
        """)
        tb = QHBoxLayout(toolbar_widget)
        tb.setContentsMargins(24, 0, 24, 0)
        tb.setSpacing(20)

        # Sección Page
        page_section = QWidget()
        page_section.setStyleSheet("QWidget { border: none; background: transparent; }")
        page_layout = QVBoxLayout(page_section)
        page_layout.setContentsMargins(0, 8, 0, 8)
        page_layout.setSpacing(4)
        
        page_label = QLabel('PAGE')
        page_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 1px;
                text-decoration: none;
                border: none;
                background: transparent;
            }}
        """)
        page_layout.addWidget(page_label)
        
        self.page_spinbox = QSpinBox()
        self.page_spinbox.setRange(1, 1)
        self.page_spinbox.valueChanged.connect(self.on_page_changed)
        self.page_spinbox.setFixedSize(80, 36)
        self.page_spinbox.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_base']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 6px 8px;
                font-size: 14px;
                font-weight: 600;
                color: {COLORS['text_primary']};
            }}
            QSpinBox:focus {{
                border-color: {COLORS['accent']};
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                width: 0px;
                height: 0px;
                border: none;
                background: transparent;
            }}
        """)
        
        # Crear botones personalizados con texto
        up_btn = QPushButton('▲')
        up_btn.setParent(self.page_spinbox)
        up_btn.setGeometry(64, 2, 16, 16)
        up_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 9px;
                text-align: right;
                padding-right: 1px;
            }}
            QPushButton:hover {{
                background: {COLORS['accent']};
                color: white;
                border-radius: 3px;
            }}
        """)
        up_btn.clicked.connect(lambda: self.page_spinbox.setValue(self.page_spinbox.value() + 1))
        
        down_btn = QPushButton('▼')
        down_btn.setParent(self.page_spinbox)
        down_btn.setGeometry(64, 18, 16, 16)
        down_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 9px;
                text-align: right;
                padding-right: 1px;
            }}
            QPushButton:hover {{
                background: {COLORS['accent']};
                color: white;
                border-radius: 3px;
            }}
        """)
        down_btn.clicked.connect(lambda: self.page_spinbox.setValue(self.page_spinbox.value() - 1))
        page_layout.addWidget(self.page_spinbox)
        tb.addWidget(page_section)

        # Separador
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"background: {COLORS['border']}; max-width: 1px; border: none;")
        sep1.setFixedHeight(40)
        tb.addWidget(sep1)

        # Sección Mode
        mode_section = QWidget()
        mode_section.setStyleSheet("QWidget { border: none; background: transparent; }")
        mode_section_layout = QVBoxLayout(mode_section)
        mode_section_layout.setContentsMargins(0, 8, 0, 8)
        mode_section_layout.setSpacing(4)
        
        mode_label = QLabel('MODE')
        mode_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 1px;
                text-decoration: none;
                border: none;
                background: transparent;
            }}
        """)
        mode_section_layout.addWidget(mode_label)

        # Mode buttons mejorados
        mode_frame = QWidget()
        mode_frame.setFixedHeight(36)
        mode_frame.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_base']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
            }}
        """)
        mode_lay = QHBoxLayout(mode_frame)
        mode_lay.setContentsMargins(4, 4, 4, 4)
        mode_lay.setSpacing(2)

        self.mode_group = QButtonGroup(self)
        self.col_radio = QRadioButton('Columns')
        self.col_radio.setChecked(True)
        self.col_radio.toggled.connect(lambda c: self.set_mode('column') if c else None)
        self.col_radio.setFixedHeight(28)
        self.col_radio.setStyleSheet(f"""
            QRadioButton {{
                color: {COLORS['text_secondary']};
                font-size: 12px;
                font-weight: 600;
                padding: 0 16px;
                border-radius: 6px;
                background: transparent;
            }}
            QRadioButton:checked {{
                background: {COLORS['accent']};
                color: white;
            }}
            QRadioButton:hover:!checked {{
                background: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
            }}
            QRadioButton::indicator {{ width: 0; height: 0; }}
        """)
        self.mode_group.addButton(self.col_radio)
        mode_lay.addWidget(self.col_radio)

        self.row_radio = QRadioButton('Rows')
        self.row_radio.toggled.connect(lambda c: self.set_mode('row') if c else None)
        self.row_radio.setFixedHeight(28)
        self.row_radio.setStyleSheet(self.col_radio.styleSheet())
        self.mode_group.addButton(self.row_radio)
        mode_lay.addWidget(self.row_radio)
        
        mode_section_layout.addWidget(mode_frame)
        tb.addWidget(mode_section)

        # Separador
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"background: {COLORS['border']}; max-width: 1px; border: none;")
        sep2.setFixedHeight(40)
        tb.addWidget(sep2)

        # Sección Zoom
        zoom_section = QWidget()
        zoom_section.setStyleSheet("QWidget { border: none; background: transparent; }")
        zoom_layout = QVBoxLayout(zoom_section)
        zoom_layout.setContentsMargins(0, 8, 0, 8)
        zoom_layout.setSpacing(4)
        
        zoom_label = QLabel('ZOOM')
        zoom_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 1px;
                text-decoration: none;
                border: none;
                background: transparent;
            }}
        """)
        zoom_layout.addWidget(zoom_label)
        
        zoom_container = QWidget()
        zoom_container.setStyleSheet("QWidget { border: none; background: transparent; }")
        zoom_container_layout = QHBoxLayout(zoom_container)
        zoom_container_layout.setContentsMargins(0, 0, 0, 0)
        zoom_container_layout.setSpacing(12)
        
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(25, 200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setFixedWidth(120)
        self.zoom_slider.setFixedHeight(36)
        self.zoom_slider.valueChanged.connect(self.on_zoom_changed)
        self.zoom_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                border: none;
                height: 6px;
                background: {COLORS['bg_base']};
                border-radius: 3px;
            }}
            QSlider::handle:horizontal {{
                background: {COLORS['accent']};
                border: none;
                width: 20px;
                height: 20px;
                margin: -7px 0;
                border-radius: 10px;
            }}
            QSlider::handle:horizontal:hover {{
                background: {COLORS['accent_hover']};
            }}
            QSlider::sub-page:horizontal {{
                background: {COLORS['accent']};
                border-radius: 3px;
            }}
        """)
        zoom_container_layout.addWidget(self.zoom_slider)
        
        self.zoom_label = QLabel('100%')
        self.zoom_label.setFixedWidth(45)
        self.zoom_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 12px;
                font-weight: 600;
                background: {COLORS['bg_base']};
                border: 2px solid {COLORS['border']};
                border-radius: 6px;
                padding: 6px 8px;
            }}
        """)
        self.zoom_label.setAlignment(Qt.AlignCenter)
        zoom_container_layout.addWidget(self.zoom_label)
        
        zoom_layout.addWidget(zoom_container)
        tb.addWidget(zoom_section)

        tb.addStretch()

        # Sección Status - Count badges mejorados
        status_section = QWidget()
        status_section.setStyleSheet("QWidget { border: none; background: transparent; }")
        status_layout = QVBoxLayout(status_section)
        status_layout.setContentsMargins(0, 8, 0, 8)
        status_layout.setSpacing(4)
        
        status_label = QLabel('STATUS')
        status_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 1px;
                text-decoration: none;
                border: none;
                background: transparent;
            }}
        """)
        status_layout.addWidget(status_label)
        
        badges_container = QWidget()
        badges_container.setStyleSheet("QWidget { border: none; background: transparent; }")
        badges_layout = QHBoxLayout(badges_container)
        badges_layout.setContentsMargins(0, 0, 0, 0)
        badges_layout.setSpacing(8)
        
        self.cols_badge = QLabel('0 cols')
        self.cols_badge.setFixedHeight(28)
        self.cols_badge.setStyleSheet(f"""
            QLabel {{
                background: {COLORS['accent_dim']};
                color: {COLORS['accent']};
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 11px;
                font-weight: 700;
            }}
        """)
        badges_layout.addWidget(self.cols_badge)
        
        self.rows_badge = QLabel('0 rows')
        self.rows_badge.setFixedHeight(28)
        self.rows_badge.setStyleSheet(f"""
            QLabel {{
                background: {COLORS['accent_dim']};
                color: {COLORS['accent']};
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 11px;
                font-weight: 700;
            }}
        """)
        badges_layout.addWidget(self.rows_badge)
        
        status_layout.addWidget(badges_container)
        tb.addWidget(status_section)

        # Separador
        sep3 = QFrame()
        sep3.setFrameShape(QFrame.VLine)
        sep3.setStyleSheet(f"background: {COLORS['border']}; max-width: 1px; border: none;")
        sep3.setFixedHeight(40)
        tb.addWidget(sep3)

        # Sección Actions - Botones mejorados
        actions_section = QWidget()
        actions_section.setStyleSheet("QWidget { border: none; background: transparent; }")
        actions_layout = QVBoxLayout(actions_section)
        actions_layout.setContentsMargins(0, 8, 0, 8)
        actions_layout.setSpacing(4)
        
        actions_label = QLabel('ACTIONS')
        actions_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 1px;
                text-decoration: none;
                border: none;
                background: transparent;
            }}
        """)
        actions_layout.addWidget(actions_label)
        
        buttons_container = QWidget()
        buttons_container.setStyleSheet("QWidget { border: none; background: transparent; }")
        buttons_layout = QHBoxLayout(buttons_container)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(8)
        
        clear_c = QPushButton('Clear Cols')
        clear_c.setFixedHeight(28)
        clear_c.clicked.connect(self.clear_columns)
        clear_c.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['bg_base']};
                color: {COLORS['text_secondary']};
                border: 2px solid {COLORS['border']};
                border-radius: 6px;
                padding: 0 12px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
                border-color: {COLORS['accent']};
            }}
            QPushButton:pressed {{
                background: {COLORS['accent_dim']};
            }}
        """)
        buttons_layout.addWidget(clear_c)
        
        clear_r = QPushButton('Clear Rows')
        clear_r.setFixedHeight(28)
        clear_r.clicked.connect(self.clear_rows)
        clear_r.setStyleSheet(clear_c.styleSheet())
        buttons_layout.addWidget(clear_r)
        
        actions_layout.addWidget(buttons_container)
        tb.addWidget(actions_section)

        lay.addWidget(toolbar_widget)

        # ── Canvas ────────────────────────────
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.Antialiasing)
        self.view.setRenderHint(QPainter.SmoothPixmapTransform)
        self.view.setDragMode(QGraphicsView.NoDrag)  # Cambiar a NoDrag para manejar manualmente
        self.view.setMouseTracking(True)
        self.view.viewport().installEventFilter(self)
        self.view.setStyleSheet(f"background: {COLORS['bg_base']}; border: none;")
        
        # Variables para pan
        self.pan_active = False
        self.last_pan_point = None
        
        # Instalar event filter personalizado para zoom y pan
        self.view.wheelEvent = self.view_wheel_event
        self.view.mousePressEvent = self.view_mouse_press_event
        self.view.mouseMoveEvent = self.view_mouse_move_event
        self.view.mouseReleaseEvent = self.view_mouse_release_event
        
        # Habilitar el seguimiento del botón medio
        self.view.setContextMenuPolicy(Qt.NoContextMenu)
        
        lay.addWidget(self.view, 1)

        # ── Footer ────────────────────────────
        footer = QWidget()
        footer.setFixedHeight(56)
        footer.setStyleSheet(f"""
            background: {COLORS['bg_surface']};
            border-top: 1px solid {COLORS['border']};
        """)
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(16, 0, 16, 0)
        fl.setSpacing(12)

        self.info_label = QLabel('Click to add lines  ·  Right-click to remove nearest')
        self.info_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 12px;")
        fl.addWidget(self.info_label)
        fl.addStretch()
        
        # Botón para guardar como plantilla
        save_template = make_btn('💾 Save as Template', 'purple', 'sm')
        save_template.clicked.connect(self.save_as_template)
        fl.addWidget(save_template)

        cancel = make_btn('Cancel', 'ghost')
        cancel.clicked.connect(self.reject)
        fl.addWidget(cancel)

        save = make_btn('Apply Grid', 'success')
        save.clicked.connect(self.accept)
        fl.addWidget(save)

        lay.addWidget(footer)
        self.preview_line = None
    
    def title_bar_mouse_press(self, event):
        """Manejar clic en la barra de título para iniciar arrastre"""
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
    
    def title_bar_mouse_move(self, event):
        """Manejar movimiento del mouse para arrastrar la ventana"""
        if event.buttons() == Qt.LeftButton and self.drag_position is not None:
            self.move(event.globalPos() - self.drag_position)
            event.accept()
    
    def view_wheel_event(self, event):
        """Manejar zoom con rueda del ratón"""
        # Obtener el factor de zoom
        zoom_in_factor = 1.15
        zoom_out_factor = 1 / zoom_in_factor
        
        # Guardar la posición del cursor en la escena
        old_pos = self.view.mapToScene(event.pos())
        
        # Aplicar zoom
        if event.angleDelta().y() > 0:
            zoom_factor = zoom_in_factor
            new_zoom = min(200, self.zoom_slider.value() * zoom_in_factor)
        else:
            zoom_factor = zoom_out_factor
            new_zoom = max(25, self.zoom_slider.value() * zoom_out_factor)
        
        # Actualizar el slider (esto llamará automáticamente a on_zoom_changed)
        self.zoom_slider.setValue(int(new_zoom))
        
        # Mantener la posición del cursor en la misma ubicación de la escena
        new_pos = self.view.mapToScene(event.pos())
        delta = new_pos - old_pos
        self.view.translate(delta.x(), delta.y())
        
        event.accept()
    
    def view_mouse_press_event(self, event):
        """Manejar clic del ratón en el view"""
        print(f"Mouse press: button={event.button()}, pos={event.pos()}")  # Debug
        
        if event.button() == Qt.MiddleButton:
            # Iniciar pan con botón medio
            print("Middle button pressed - starting pan")  # Debug
            self.pan_active = True
            self.last_pan_point = event.pos()
            self.view.setCursor(Qt.ClosedHandCursor)
            event.accept()
        elif event.button() == Qt.LeftButton and event.modifiers() & Qt.ControlModifier:
            # Iniciar pan con Ctrl+clic izquierdo
            print("Ctrl+Left button pressed - starting pan")  # Debug
            self.pan_active = True
            self.last_pan_point = event.pos()
            self.view.setCursor(Qt.ClosedHandCursor)
            event.accept()
        elif event.button() == Qt.LeftButton:
            # Comportamiento normal para añadir líneas
            self.on_mouse_press(event)
        else:
            # Llamar al comportamiento original para otros botones
            QGraphicsView.mousePressEvent(self.view, event)
    
    def view_mouse_move_event(self, event):
        """Manejar movimiento del ratón en el view"""
        if self.pan_active and self.last_pan_point is not None:
            # Realizar pan
            delta = event.pos() - self.last_pan_point
            self.last_pan_point = event.pos()
            
            # Mover la vista usando las barras de desplazamiento
            h_bar = self.view.horizontalScrollBar()
            v_bar = self.view.verticalScrollBar()
            h_bar.setValue(h_bar.value() - delta.x())
            v_bar.setValue(v_bar.value() - delta.y())
            
            print(f"Panning: delta=({delta.x()}, {delta.y()})")  # Debug
            event.accept()
        else:
            # Comportamiento normal para preview de líneas
            self.on_mouse_move(event)
    
    def view_mouse_release_event(self, event):
        """Manejar liberación del ratón en el view"""
        if event.button() == Qt.MiddleButton or (event.button() == Qt.LeftButton and self.pan_active):
            # Terminar pan
            print("Ending pan")  # Debug
            self.pan_active = False
            self.last_pan_point = None
            self.view.setCursor(Qt.ArrowCursor)
            event.accept()
        else:
            # Llamar al comportamiento original
            QGraphicsView.mouseReleaseEvent(self.view, event)

    def eventFilter(self, obj, event):
        from PyQt5.QtCore import QEvent
        if obj == self.view.viewport():
            if event.type() == QEvent.MouseButtonPress:
                self.on_mouse_press(event); return True
            elif event.type() == QEvent.MouseMove:
                self.on_mouse_move(event); return True
        return super().eventFilter(obj, event)

    def on_mouse_press(self, event):
        """Manejar clic del mouse para añadir líneas de grilla"""
        if not self.pdf_doc: 
            return
            
        # Convertir coordenadas del mouse a coordenadas de la escena
        scene_pos = self.view.mapToScene(event.pos())
        x, y = scene_pos.x(), scene_pos.y()
        
        # Verificar que estamos dentro de los límites de la página
        if x < 0 or x > self.page_width or y < 0 or y > self.page_height: 
            return
            
        if event.button() == Qt.LeftButton:
            if self.current_mode == 'column':
                self.column_lines.append(x)
                self.column_lines.sort()
            else:
                self.row_lines.append(y)
                self.row_lines.sort()
            self.update_lines()
            
        elif event.button() == Qt.RightButton:
            if self.current_mode == 'column' and self.column_lines:
                closest_line = min(self.column_lines, key=lambda lx: abs(lx - x))
                self.column_lines.remove(closest_line)
            elif self.current_mode == 'row' and self.row_lines:
                closest_line = min(self.row_lines, key=lambda ly: abs(ly - y))
                self.row_lines.remove(closest_line)
            self.update_lines()

    def on_mouse_move(self, event):
        """Manejar movimiento del mouse para mostrar vista previa de líneas"""
        if not self.pdf_doc: 
            return
            
        # Convertir coordenadas del mouse a coordenadas de la escena
        scene_pos = self.view.mapToScene(event.pos())
        x, y = scene_pos.x(), scene_pos.y()
        
        # Remover línea de vista previa anterior
        if self.preview_line:
            self.scene.removeItem(self.preview_line)
            self.preview_line = None
            
        # Crear nueva línea de vista previa
        pen = QPen(QColor(59, 130, 246, 80), 2.0, Qt.DashLine)
        if self.current_mode == 'column':
            self.preview_line = self.scene.addLine(x, 0, x, self.page_height, pen)
        else:
            self.preview_line = self.scene.addLine(0, y, self.page_width, y, pen)

    def on_load_pdf(self):
        fp, _ = QFileDialog.getOpenFileName(self, 'Select PDF', '', 'PDF Files (*.pdf)')
        if fp: self.load_pdf(fp)

    def load_pdf(self, path):
        try:
            if self.pdf_doc: self.pdf_doc.close()
            self.pdf_path = path
            self.pdf_doc = fitz.open(path)
            self.page_spinbox.setMaximum(len(self.pdf_doc))
            self.page_spinbox.setValue(1)
            self.render_page()
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

    def on_page_changed(self, v):
        self.page_num = v - 1; self.render_page()

    def on_zoom_changed(self, v):
        """Manejar cambio de zoom desde el slider"""
        self.zoom_factor = v / 100.0
        self.zoom_label.setText(f'{v}%')
        
        # Aplicar transformación de zoom al view
        transform = QTransform()
        transform.scale(self.zoom_factor, self.zoom_factor)
        self.view.setTransform(transform)
        
        # Actualizar las líneas con el nuevo factor de zoom
        self.update_lines()

    def set_mode(self, mode):
        self.current_mode = mode

    def render_page(self):
        """Renderizar la página del PDF"""
        if not self.pdf_doc: 
            return
            
        page = self.pdf_doc[self.page_num]
        self.page_width, self.page_height = page.rect.width, page.rect.height
        
        # Renderizar a resolución base (sin zoom aplicado aquí)
        mat = fitz.Matrix(2, 2)  # Resolución fija para calidad
        pix = page.get_pixmap(matrix=mat)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        
        # Crear pixmap a tamaño base
        pixmap = QPixmap.fromImage(img).scaled(
            int(self.page_width), int(self.page_height),
            Qt.KeepAspectRatio, Qt.SmoothTransformation)
        
        # Limpiar escena y añadir pixmap
        self.scene.clear()
        self.preview_line = None
        self.scene.addPixmap(pixmap)
        self.scene.setSceneRect(0, 0, self.page_width, self.page_height)
        
        # Actualizar líneas
        self.update_lines()

    def update_lines(self):
        """Actualizar las líneas de la grilla"""
        # Remover líneas existentes
        for item in self.scene.items():
            if isinstance(item, QGraphicsLineItem): 
                self.scene.removeItem(item)
        
        # Añadir líneas de columnas (azules)
        for x in self.column_lines:
            pen = QPen(QColor(59, 130, 246), 2.0)  # Líneas más gruesas para mejor visibilidad
            self.scene.addLine(x, 0, x, self.page_height, pen)
        
        # Añadir líneas de filas (verdes)
        for y in self.row_lines:
            pen = QPen(QColor(16, 185, 129), 2.0)  # Líneas más gruesas para mejor visibilidad
            self.scene.addLine(0, y, self.page_width, y, pen)
        
        # Actualizar badges
        self.cols_badge.setText(f'{len(self.column_lines)} cols')
        self.rows_badge.setText(f'{len(self.row_lines)} rows')
        
        # Guardar configuración
        self.save_config()

    def clear_columns(self): self.column_lines.clear(); self.update_lines()
    def clear_rows(self): self.row_lines.clear(); self.update_lines()
    
    def save_as_template(self):
        """Guardar la configuración actual como plantilla"""
        if len(self.column_lines) < 2 or len(self.row_lines) < 2:
            QMessageBox.warning(self, 'Insufficient Grid Lines', 
                              'Need at least 2 column lines and 2 row lines to save as template.')
            return
        
        # Pedir nombre para la plantilla
        from PyQt5.QtWidgets import QInputDialog
        nc = len(self.column_lines) - 1
        nr = len(self.row_lines) - 1
        name, ok = QInputDialog.getText(self, 'Save Grid Template', 
                                        'Enter a name for this grid template:',
                                        text=f'Grid {nc}×{nr}')
        if not ok or not name.strip():
            return
        
        name = name.strip()
        
        # Cargar plantillas existentes
        templates_file = os.path.join(get_app_path(), 'grid_templates.json')
        templates = {}
        if os.path.exists(templates_file):
            try:
                with open(templates_file, 'r') as f:
                    templates = json.load(f)
            except Exception as e:
                print(f'Error loading templates: {e}')
        
        # Verificar si ya existe
        if name in templates:
            reply = QMessageBox.question(self, 'Template Exists', 
                                        f'A template named "{name}" already exists. Overwrite?',
                                        QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        
        # Guardar plantilla
        templates[name] = {
            'cols': nc,
            'rows': nr,
            'column_positions': sorted(self.column_lines),
            'row_positions': sorted(self.row_lines),
            'page_width': self.page_width,
            'page_height': self.page_height,
            'col_sizes': '',
            'row_sizes': ''
        }
        
        try:
            with open(templates_file, 'w') as f:
                json.dump(templates, f, indent=2)
            QMessageBox.information(self, 'Template Saved', 
                                   f'Grid template "{name}" saved successfully!')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error saving template: {e}')

    def get_grid_data(self):
        return {
            'column_positions': sorted(self.column_lines),
            'row_positions': sorted(self.row_lines),
            'page_width': self.page_width,
            'page_height': self.page_height
        }

    def get_config_file_path(self):
        return os.path.join(get_app_path(), 'grid_config.json')

    def save_config(self):
        try:
            with open(self.get_config_file_path(), 'w', encoding='utf-8') as f:
                json.dump({
                    'column_lines': sorted(self.column_lines),
                    'row_lines': sorted(self.row_lines),
                    'page_width': self.page_width,
                    'page_height': self.page_height,
                    'page_num': self.page_num,
                    'zoom_factor': self.zoom_factor
                }, f, indent=2)
        except Exception as e:
            print(f'Config save error: {e}')

    def load_saved_config(self):
        cp = self.get_config_file_path()
        if not os.path.exists(cp): return False
        try:
            with open(cp, 'r', encoding='utf-8') as f:
                d = json.load(f)
            self.column_lines = d.get('column_lines', [])
            self.row_lines = d.get('row_lines', [])
            self.zoom_slider.setValue(int(d.get('zoom_factor', 1.0) * 100))
            self.update_lines()
            return True
        except Exception as e:
            print(f'Config load error: {e}'); return False

    def closeEvent(self, event):
        self.save_config()
        if self.pdf_doc: self.pdf_doc.close()
        event.accept()


# ─────────────────────────────────────────────
#  MAIN WINDOW
# ─────────────────────────────────────────────
class PDFReferenceDetector(QMainWindow):

    REFERENCE_PATTERNS = {
        'Style /1.0-A': {
            'pattern': r'/\s*(\d+)[.\s]+(\d+|[A-Za-z]+)\s*[-/]\s*([A-Za-z0-9]+)',
            'example': '/1.0-A, /10.5-Z',
            'groups': ('página', 'columna', 'fila'),
            'order': 'page.col-row'
        },
        'Style 25-A.0': {
            'pattern': r'(\d+)\s*[-]\s*([A-Za-z]+)[.\s]+(\d+)',
            'example': '25-A.0, 10-B.5',
            'groups': ('página', 'fila', 'columna'),
            'order': 'page-row.col'
        },
        'Style A1/25': {
            'pattern': r'([A-Za-z]+)(\d+)\s*[/]\s*(\d+)',
            'example': 'A1/25, B5/10',
            'groups': ('fila', 'columna', 'página'),
            'order': 'row+col/page'
        },
        'Style (1-A-0)': {
            'pattern': r'\(\s*(\d+)\s*[-]\s*([A-Za-z]+)\s*[-]\s*(\d+)\s*\)',
            'example': '(1-A-0), (10-B-5)',
            'groups': ('página', 'fila', 'columna'),
            'order': '(page-row-col)'
        },
        'Custom': {
            'pattern': '',
            'example': 'Define your own regex',
            'groups': ('group1', 'group2', 'group3'),
            'order': 'custom'
        }
    }

    def __init__(self):
        super().__init__()
        self.pdf_path = None
        self.pdf_paths = []
        self.references = []
        self.all_references = {}
        self.pdf_document = None
        self.current_pattern = 'Style /1.0-A'
        self.custom_pattern = ''
        self.column_positions = []
        self.row_positions = []
        self.grid_detected = False
        self.init_ui()
        self.load_saved_grid_config()

    def closeEvent(self, event):
        if self.pdf_document: self.pdf_document.close()
        event.accept()

    # ─── STYLESHEET HELPERS ───────────────────

    def init_ui(self):
        self.setWindowTitle('PDF Reference Detector')
        
        # 1. Quitar completamente el control de Windows
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.NoDropShadowWindowHint  # Quita también la sombra del SO
        )
        
        # 2. Fondo transparente para que no salgan picos
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        self.setGeometry(60, 60, 1320, 860)
        self.setMinimumSize(900, 640)
        
        # Centrar la ventana en el monitor
        self.center_on_screen()

        icon_path = os.path.join(get_app_path(), 'logo.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setAcceptDrops(True)
        
        # Radio de redondeo para la ventana (más redondeado)
        self.border_radius = 35
        
        # Variables para arrastrar la ventana
        self.drag_position = None

        self.setup_ui()
    
    def center_on_screen(self):
        """Centrar la ventana en el monitor principal"""
        from PyQt5.QtWidgets import QApplication
        screen = QApplication.desktop().screenGeometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

    def showEvent(self, event):
        """Se ejecuta cuando la ventana se muestra por primera vez"""
        super().showEvent(event)
        hwnd = int(self.winId())
        self._quitar_control_windows(hwnd)
        self._aplicar_mascara_redondeada()

    def _aplicar_mascara_redondeada(self):
        """Aplica una máscara redondeada para eliminar completamente los picos"""
        try:
            # Crear una región redondeada que coincida exactamente con el paintEvent
            path = QPainterPath()
            path.addRoundedRect(QRectF(1, 1, self.width() - 2, self.height() - 2), 
                               self.border_radius, self.border_radius)
            
            # Convertir el path a una región y aplicarla como máscara
            region = QRegion(path.toFillPolygon().toPolygon())
            self.setMask(region)
        except Exception as e:
            print(f"No se pudo aplicar la máscara: {e}")

    def resizeEvent(self, event):
        """Actualizar la máscara cuando la ventana cambie de tamaño"""
        super().resizeEvent(event)
        if self.isVisible():
            self._aplicar_mascara_redondeada()

    def _quitar_control_windows(self, hwnd):
        """Quita completamente el control de Windows sobre la ventana"""
        try:
            dwmapi = ctypes.windll.dwmapi
            
            # 1. Quitar redondeo forzado de Windows 11
            DWMWA_WINDOW_CORNER_PREFERENCE = 33
            DWMWCP_DONOTROUND = 1  # Sin redondeo del sistema
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 
                DWMWA_WINDOW_CORNER_PREFERENCE,
                ctypes.byref(ctypes.c_int(DWMWCP_DONOTROUND)), 
                ctypes.sizeof(ctypes.c_int)
            )
            
            # 2. Extender marco DWM (elimina sombras y decoración)
            class MARGINS(ctypes.Structure):
                _fields_ = [("left", ctypes.c_int), ("right", ctypes.c_int),
                           ("top", ctypes.c_int), ("bottom", ctypes.c_int)]
            
            margins = MARGINS(-1, -1, -1, -1)  # -1 = extender a toda la ventana
            dwmapi.DwmExtendFrameIntoClientArea(hwnd, ctypes.byref(margins))
            
        except Exception as e:
            # Si falla (por ejemplo, en versiones antiguas de Windows), continuar normalmente
            print(f"No se pudo quitar el control de Windows: {e}")

    def paintEvent(self, event):
        """Dibuja la ventana con borde naranja simple y directo"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Dibujar el fondo blanco primero
        bg_path = QPainterPath()
        bg_path.addRoundedRect(QRectF(0, 0, self.width(), self.height()), 
                              self.border_radius, self.border_radius)
        painter.fillPath(bg_path, QColor('#FFFFFF'))
        
        # Dibujar el borde naranja directamente
        painter.setPen(QPen(QColor('#FF6B2B'), 6, Qt.SolidLine))
        painter.setBrush(Qt.NoBrush)  # Sin relleno, solo borde
        
        # Dibujar el rectángulo redondeado con borde
        border_rect = QRectF(3, 3, self.width() - 6, self.height() - 6)
        painter.drawRoundedRect(border_rect, self.border_radius - 3, self.border_radius - 3)
        
        painter.end()

    def setup_ui(self):
        # Aplicar estilo personalizado
        custom_style = APP_STYLESHEET + f"""
        QMainWindow {{
            background: transparent;
        }}
        
        /* Bordes mucho más redondeados para componentes específicos */
        QPushButton {{
            border-radius: 24px !important;
        }}
        
        QSpinBox {{
            border-radius: 20px !important;
        }}
        
        QTableWidget {{
            border-radius: 30px !important;
        }}
        
        QTextEdit {{
            border-radius: 30px !important;
        }}
        
        QListWidget {{
            border-radius: 30px !important;
        }}
        """
        self.setStyleSheet(custom_style)

        root = QWidget()
        root.setStyleSheet(f"background: {COLORS['bg_base']};")
        self.setCentralWidget(root)
        root_lay = QVBoxLayout(root)
        root_lay.setSpacing(0)
        root_lay.setContentsMargins(6, 6, 6, 6)  # Margen para el borde de 6px

        # ── CUSTOM TITLE BAR ──────────────────
        title_bar = self._create_title_bar()
        root_lay.addWidget(title_bar)

        # ── MAIN CONTENT CONTAINER ───────────
        content_container = QWidget()
        content_lay = QHBoxLayout(content_container)
        content_lay.setSpacing(0)
        content_lay.setContentsMargins(0, 0, 0, 0)

        # ── LEFT SIDEBAR ──────────────────────
        sidebar = self._build_sidebar()
        content_lay.addWidget(sidebar)

        # ── MAIN CONTENT ──────────────────────
        content = self._build_content()
        content_lay.addWidget(content, 1)

        root_lay.addWidget(content_container, 1)

        # Status bar con barra de progreso integrada
        sb = self.statusBar()
        sb.showMessage('Ready  —  Drag PDFs or click  Select Files')
        
        # Crear barra de progreso integrada
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)  # Oculta inicialmente
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                background: {COLORS['bg_elevated']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                text-align: center;
                color: {COLORS['text_primary']};
                font-weight: 600;
                font-size: 10px;
            }}
            QProgressBar::chunk {{
                background: #FF6B2B;
                border-radius: 6px;
                margin: 1px;
            }}
        """)
        
        # Agregar la barra de progreso a la status bar
        sb.addPermanentWidget(self.progress_bar)
        
        sb.setStyleSheet(f"""
            QStatusBar {{
                background: {COLORS['bg_surface']};
                border-bottom-left-radius: 35px;
                border-bottom-right-radius: 35px;
                color: {COLORS['text_muted']};
                text-align: center;
            }}
            QStatusBar::item {{
                border: none;
            }}
        """)

        # Variables para arrastrar la ventana
        self.drag_position = None

        # Compat stubs
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.ref_count_label = QLabel('0')
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)

    def mousePressEvent(self, event):
        """Manejar clic para arrastrar la ventana"""
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        """Manejar movimiento del mouse para arrastrar la ventana"""
        if event.buttons() == Qt.LeftButton and self.drag_position:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    # ─── SIDEBAR ─────────────────────────────
    def _create_title_bar(self):
        """Crear barra de título personalizada"""
        title_bar = QWidget()
        title_bar.setFixedHeight(40)
        title_bar.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_surface']};
                border-top-left-radius: 35px;
                border-top-right-radius: 35px;
            }}
        """)
        
        layout = QHBoxLayout(title_bar)
        layout.setContentsMargins(16, 0, 32, 0)
        layout.setSpacing(12)
        
        # Espacio izquierdo (compensar por los botones de la derecha)
        layout.addStretch(2)
        
        # Título centrado
        title_label = QLabel('PDF Reference Detector')
        title_label.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 14px;
            font-weight: 600;
        """)
        layout.addWidget(title_label)
        
        # Espacio derecho menor para compensar los botones
        layout.addStretch(1)
        
        # Botones de control de ventana (un poco más a la izquierda)
        btn_style = f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 18px;
                padding: 6px 10px;
                border-radius: 8px;
            }}
            QPushButton:hover {{
                background: {COLORS['bg_hover']};
            }}
        """
        
        # Reducir espacio antes de los botones para moverlos un poco más a la izquierda
        layout.addSpacing(34)
        
        # Minimizar
        min_btn = QPushButton('−')
        min_btn.setFixedSize(36, 28)
        min_btn.setStyleSheet(btn_style)
        min_btn.clicked.connect(self.showMinimized)
        layout.addWidget(min_btn)
        
        # Cerrar
        close_btn = QPushButton('×')
        close_btn.setFixedSize(36, 28)
        close_btn.setStyleSheet(btn_style + f"""
            QPushButton:hover {{
                background: {COLORS['danger']};
                color: white;
            }}
        """)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        return title_bar

    def title_bar_double_click(self, event):
        """Manejar doble clic en la barra de título para maximizar/restaurar"""
        if event.button() == Qt.LeftButton:
            if self.isMaximized():
                self.showNormal()
            else:
                self.showMaximized()
            event.accept()

    def title_bar_mouse_press(self, event):
        """Manejar clic en la barra de título para arrastrar"""
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def title_bar_mouse_move(self, event):
        """Manejar movimiento del mouse para arrastrar la ventana"""
        if event.buttons() == Qt.LeftButton and self.drag_position:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def _build_sidebar(self):
        w = QWidget()
        w.setFixedWidth(220)
        w.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_surface']};
            }}
        """)
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Navigation section (movido arriba, sin brand)
        nav_area = QWidget()
        nav_area.setStyleSheet(f"background: {COLORS['bg_surface']};")
        nav_lay = QVBoxLayout(nav_area)
        nav_lay.setContentsMargins(12, 20, 12, 16)  # Más padding arriba
        nav_lay.setSpacing(4)

        nav_lay.addWidget(section_label('Navigation'))
        nav_lay.addSpacing(8)

        nav_items = [
            ('Files',    'files_icon',    True),
            ('Pattern',  'pattern_icon',  False),
            ('Grid',     'grid_icon',     False),
            ('Style',    'style_icon',    False),
        ]
        self._nav_btns = []
        for label, icon_key, active in nav_items:
            # Crear botón con layout horizontal para icono SVG + texto
            btn = QPushButton()
            btn.setCheckable(True)
            btn.setChecked(active)
            btn.setFixedHeight(38)
            btn.setObjectName('navbtn')
            btn.setStyleSheet(self._nav_btn_style())
            btn.clicked.connect(lambda _, i=len(self._nav_btns): self._switch_section(i))
            
            # Crear layout horizontal para el botón
            btn_layout = QHBoxLayout(btn)
            btn_layout.setContentsMargins(12, 0, 12, 0)
            btn_layout.setSpacing(8)
            
            # Agregar icono SVG más grande
            if icon_key in SVG_ICONS:
                icon_widget = create_svg_icon(SVG_ICONS[icon_key], 24)  # Aumentado de 18 a 24
            else:
                # Fallback a emoji si no hay SVG
                fallback_icons = {'files_icon': '📁', 'pattern_icon': '🔤', 'grid_icon': '📐', 'style_icon': '🎨'}
                icon_widget = QLabel(fallback_icons.get(icon_key, '•'))
                icon_widget.setStyleSheet("font-size: 20px; background: transparent; border: none;")  # Aumentado de 16px a 20px
            
            btn_layout.addWidget(icon_widget)
            
            # Agregar texto
            text_label = QLabel(label)
            text_label.setStyleSheet("background: transparent; border: none; color: inherit;")
            btn_layout.addWidget(text_label)
            
            btn_layout.addStretch()
            
            self._nav_btns.append(btn)
            nav_lay.addWidget(btn)

        nav_lay.addSpacing(24)

        nav_lay.addStretch()

        # Grid template selector button at bottom
        nav_lay.addSpacing(16)
        self.grid_template_btn = QPushButton('○  No grid configured')
        self.grid_template_btn.setMinimumHeight(36)  # Altura mínima para evitar cortes
        self.grid_template_btn.setStyleSheet(f"""
            QPushButton {{
                color: {COLORS['text_muted']};
                font-size: 10px;
                padding: 8px 6px;
                background: {COLORS['bg_elevated']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                text-align: center;
            }}
            QPushButton:hover {{
                background: {COLORS['bg_hover']};
                border-color: {COLORS['accent']};
            }}
        """)
        self.grid_template_btn.clicked.connect(self.show_grid_template_menu)
        nav_lay.addWidget(self.grid_template_btn)
        
        # Mantener config_status para compatibilidad (oculto)
        self.config_status = self.grid_template_btn

        lay.addWidget(nav_area, 1)
        return w

    def _nav_btn_style(self):
        return f"""
            QPushButton {{
                background: transparent;
                color: {COLORS['text_secondary']};
                border: none;
                border-radius: 8px;
                font-size: 13px;
                text-align: left;
                padding: 0 8px;
            }}
            QPushButton:checked {{
                background: {COLORS['accent_dim']};
                color: {COLORS['accent_hover']};
                font-weight: 600;
            }}
            QPushButton:hover:!checked {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_primary']};
            }}
        """

    def _action_btn_style(self, color):
        return f"""
            QPushButton {{
                background: transparent;
                color: {COLORS['text_secondary']};
                border: none;
                border-radius: 8px;
                font-size: 13px;
                text-align: left;
                padding: 0 8px;
            }}
            QPushButton:enabled:hover {{
                background: {COLORS['bg_elevated']};
                color: {color};
            }}
            QPushButton:disabled {{
                color: {COLORS['text_muted']};
            }}
        """

    def _switch_section(self, index):
        for i, btn in enumerate(self._nav_btns):
            btn.setChecked(i == index)
        self.stack.setCurrentIndex(index)

    # ─── MAIN CONTENT ─────────────────────────
    def _build_content(self):
        w = QWidget()
        w.setStyleSheet(f"background: {COLORS['bg_base']};")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Top bar
        topbar = self._build_topbar()
        lay.addWidget(topbar)

        # Stacked content
        self.stack = QStackedWidget()
        self.stack.setStyleSheet(f"background: {COLORS['bg_base']};")
        self.stack.addWidget(self._page_files())      # 0
        self.stack.addWidget(self._page_pattern())    # 1
        self.stack.addWidget(self._page_grid())       # 2
        self.stack.addWidget(self._page_style())      # 3
        lay.addWidget(self.stack, 1)
        return w

    def _build_topbar(self):
        bar = QWidget()
        bar.setFixedHeight(58)
        bar.setStyleSheet(f"""
            QWidget {{
                background: {COLORS['bg_surface']};
            }}
        """)
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(24, 16, 24, 16)
        bl.setSpacing(12)

        self.page_title = QLabel('Files')
        self.page_title.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 18px;
            font-weight: 700;
            background: transparent;
        """)
        bl.addWidget(self.page_title)
        bl.addStretch()

        # Status badges
        self.pdf_count_badge = Badge('0 files', 'muted')
        bl.addWidget(self.pdf_count_badge)

        self.ref_badge = Badge('0 refs', 'muted')
        bl.addWidget(self.ref_badge)
        return bar

    # ─── PAGE: FILES ──────────────────────────
    def _page_files(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(24, 24, 24, 24)
        lay.setSpacing(20)

        # Container que alterna entre drop zone y lista de archivos
        self.files_container = QWidget()
        self.files_layout = QVBoxLayout(self.files_container)
        self.files_layout.setContentsMargins(0, 0, 0, 0)
        self.files_layout.setSpacing(20)
        
        # ===== DROP ZONE (mostrado inicialmente) =====
        self.drop_zone_widget = QWidget()
        drop_zone_layout = QVBoxLayout(self.drop_zone_widget)
        drop_zone_layout.setContentsMargins(0, 0, 0, 0)
        
        # Contenedor del drop zone que siempre permanece visible
        self.drop_zone_container = QLabel()
        self.drop_zone_container.setAlignment(Qt.AlignCenter)
        self.drop_zone_container.setMinimumHeight(300)
        self.drop_zone_container.setStyleSheet(self._drop_zone_style(False))
        self.drop_zone_container.setObjectName('dropzone')

        # Layout del contenedor que cambiará entre drop zone y lista
        self.drop_zone_inner_layout = QVBoxLayout(self.drop_zone_container)
        self.drop_zone_inner_layout.setAlignment(Qt.AlignCenter)
        self.drop_zone_inner_layout.setContentsMargins(20, 20, 20, 20)
        
        # ===== CONTENIDO DEL DROP ZONE (se oculta cuando hay archivos) =====
        self.drop_zone_content = QWidget()
        self.drop_zone_content.setStyleSheet("QWidget { background: transparent; border: none; }")
        dz_content_layout = QVBoxLayout(self.drop_zone_content)
        dz_content_layout.setAlignment(Qt.AlignCenter)
        dz_content_layout.setSpacing(12)
        dz_content_layout.setContentsMargins(0, 0, 0, 0)
        
        # Icono SVG para drop zone
        if 'drop_icon' in SVG_ICONS:
            dz_icon = create_svg_icon(SVG_ICONS['drop_icon'], 48)
        else:
            dz_icon = QLabel('📥')
            dz_icon.setStyleSheet("font-size: 48px; background: transparent; border: none;")
            dz_icon.setAlignment(Qt.AlignCenter)
        dz_content_layout.addWidget(dz_icon)
        
        dz_txt = QLabel('Drop PDF files here')
        dz_txt.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 18px; font-weight: 600; background: transparent; border: none;")
        dz_txt.setAlignment(Qt.AlignCenter)
        dz_content_layout.addWidget(dz_txt)
        
        dz_sub = QLabel('or use the button below')
        dz_sub.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 14px; background: transparent; border: none;")
        dz_sub.setAlignment(Qt.AlignCenter)
        dz_content_layout.addWidget(dz_sub)

        # ===== LISTA DE ARCHIVOS DENTRO DEL DROP ZONE =====
        self.drop_zone_file_list = QWidget()
        self.drop_zone_file_list.setStyleSheet("QWidget { background: transparent; border: none; }")
        file_list_inner_layout = QVBoxLayout(self.drop_zone_file_list)
        file_list_inner_layout.setContentsMargins(0, 0, 0, 0)
        file_list_inner_layout.setSpacing(12)
        
        # File list header
        fh = QHBoxLayout()
        fh.addWidget(section_label('Loaded Files'))
        fh.addStretch()
        self.clear_list_btn = make_btn('Clear All', 'danger', 'sm')
        self.clear_list_btn.clicked.connect(self.clear_pdf_list)
        fh.addWidget(self.clear_list_btn)
        self.remove_selected_btn = make_btn('Remove', 'ghost', 'sm')
        self.remove_selected_btn.clicked.connect(self.remove_selected_pdfs)
        fh.addWidget(self.remove_selected_btn)
        file_list_inner_layout.addLayout(fh)
        
        # Lista de PDFs
        self.pdf_list = QListWidget()
        self.pdf_list.setMinimumHeight(120)
        self.pdf_list.setMaximumHeight(200)
        self.pdf_list.setMinimumWidth(500)  # Aumentado de 400 a 500px para más espacio
        self.pdf_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        file_list_inner_layout.addWidget(self.pdf_list)
        
        # Agregar ambos contenidos al layout del drop zone
        self.drop_zone_inner_layout.addWidget(self.drop_zone_content)
        self.drop_zone_inner_layout.addWidget(self.drop_zone_file_list)
        self.drop_zone_file_list.hide()  # Ocultar inicialmente la lista
        
        drop_zone_layout.addWidget(self.drop_zone_container)
        
        # Espacio adicional antes del botón
        drop_zone_layout.addSpacing(30)
        
        # Botón Select Files debajo del drop zone
        select_btn_layout = QHBoxLayout()
        select_btn_layout.addStretch()
        self.select_button = make_btn('📂  Select Files', 'accent', 'lg')
        self.select_button.clicked.connect(self.select_pdf)
        select_btn_layout.addWidget(self.select_button)
        select_btn_layout.addStretch()
        drop_zone_layout.addLayout(select_btn_layout)
        
        self.files_layout.addWidget(self.drop_zone_widget)
        
        lay.addWidget(self.files_container)
        
        # Más espacio entre la lista de archivos y las acciones
        lay.addSpacing(40)

        # Actions row - Centrado
        actions_section = QVBoxLayout()
        actions_section.setAlignment(Qt.AlignCenter)
        
        actions_label_container = QHBoxLayout()
        actions_label_container.addStretch()
        actions_label_container.addWidget(section_label('Actions'))
        actions_label_container.addStretch()
        actions_section.addLayout(actions_label_container)
        actions_section.addSpacing(12)
        
        ar = QHBoxLayout()
        ar.setSpacing(16)
        ar.addStretch()
        
        # Detect button with label
        detect_container = QVBoxLayout()
        detect_container.setSpacing(6)
        detect_label = QLabel('Detect')
        detect_label.setAlignment(Qt.AlignCenter)
        detect_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; font-weight: 600;")
        detect_container.addWidget(detect_label)
        
        # Detect button with SVG icon
        if 'lupa_icon' in SVG_ICONS:
            # Crear un widget contenedor para el icono SVG
            detect_icon_widget = create_svg_icon(SVG_ICONS['lupa_icon'], 32)
            self.detect_button = QPushButton()
            self.detect_button.setFixedSize(80, 80)
            
            # Layout para centrar el icono en el botón
            btn_layout = QVBoxLayout(self.detect_button)
            btn_layout.addWidget(detect_icon_widget)
            btn_layout.setAlignment(Qt.AlignCenter)
            btn_layout.setContentsMargins(0, 0, 0, 0)
        else:
            self.detect_button = QPushButton('🔍')
        self.detect_button.setFixedSize(80, 80)
        self.detect_button.setEnabled(False)
        self.detect_button.setToolTip('Detect references in PDFs')
        self.detect_button.setCursor(Qt.PointingHandCursor)
        self.detect_button.clicked.connect(self.detect_references)
        self.detect_button.setStyleSheet(f"""
            QPushButton {{
                background: #3B82F6;
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 32px;
            }}
            QPushButton:hover {{
                background: #60A5FA;
                cursor: pointer;
            }}
            QPushButton:pressed {{
                background: #1D4ED8;
            }}
            QPushButton:disabled {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_muted']};
            }}
        """)
        detect_container.addWidget(self.detect_button)
        ar.addLayout(detect_container)
        
        # Generate PDF button with label
        generate_container = QVBoxLayout()
        generate_container.setSpacing(6)
        generate_label = QLabel('Generate PDF')
        generate_label.setAlignment(Qt.AlignCenter)
        generate_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; font-weight: 600;")
        generate_container.addWidget(generate_label)
        
        # Generate PDF button with SVG icon
        if 'generate_icon' in SVG_ICONS:
            generate_icon_widget = create_svg_icon(SVG_ICONS['generate_icon'], 32)
            self.generate_button = QPushButton()
            self.generate_button.setFixedSize(80, 80)
            
            btn_layout = QVBoxLayout(self.generate_button)
            btn_layout.addWidget(generate_icon_widget)
            btn_layout.setAlignment(Qt.AlignCenter)
            btn_layout.setContentsMargins(0, 0, 0, 0)
        else:
            self.generate_button = QPushButton('✨')
        self.generate_button.setFixedSize(80, 80)
        self.generate_button.setEnabled(False)
        self.generate_button.setToolTip('Generate interactive PDF')
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self.generate_interactive_pdf)
        self.generate_button.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['warning']};
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 32px;
            }}
            QPushButton:hover {{
                background: #FCD34D;
                cursor: pointer;
            }}
            QPushButton:pressed {{
                background: #78350F;
            }}
            QPushButton:disabled {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_muted']};
            }}
        """)
        generate_container.addWidget(self.generate_button)
        ar.addLayout(generate_container)
        
        # Grid Editor button with label
        editor_container = QVBoxLayout()
        editor_container.setSpacing(6)
        editor_label = QLabel('Grid Editor')
        editor_label.setAlignment(Qt.AlignCenter)
        editor_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; font-weight: 600;")
        editor_container.addWidget(editor_label)
        
        # Grid Editor button with SVG icon
        if 'grid_editor' in SVG_ICONS:
            editor_icon_widget = create_svg_icon(SVG_ICONS['grid_editor'], 32)
            self.visual_editor_button = QPushButton()
            self.visual_editor_button.setFixedSize(80, 80)
            
            btn_layout = QVBoxLayout(self.visual_editor_button)
            btn_layout.addWidget(editor_icon_widget)
            btn_layout.setAlignment(Qt.AlignCenter)
            btn_layout.setContentsMargins(0, 0, 0, 0)
        else:
            self.visual_editor_button = QPushButton('🎨')
        self.visual_editor_button.setFixedSize(80, 80)
        self.visual_editor_button.setEnabled(False)
        self.visual_editor_button.setToolTip('Open visual grid editor')
        self.visual_editor_button.setCursor(Qt.PointingHandCursor)
        self.visual_editor_button.clicked.connect(self.open_visual_editor)
        self.visual_editor_button.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['purple']};
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 32px;
            }}
            QPushButton:hover {{
                background: #A78BFA;
                cursor: pointer;
            }}
            QPushButton:pressed {{
                background: #4C1D95;
            }}
            QPushButton:disabled {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_muted']};
            }}
        """)
        editor_container.addWidget(self.visual_editor_button)
        ar.addLayout(editor_container)
        
        # Stats button with label
        stats_container = QVBoxLayout()
        stats_container.setSpacing(6)
        stats_label = QLabel('Statistics')
        stats_label.setAlignment(Qt.AlignCenter)
        stats_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px; font-weight: 600;")
        stats_container.addWidget(stats_label)
        
        # Stats button with SVG icon
        if 'stats_icon' in SVG_ICONS:
            stats_icon_widget = create_svg_icon(SVG_ICONS['stats_icon'], 32)
            self.stats_button = QPushButton()
            self.stats_button.setFixedSize(80, 80)
            
            btn_layout = QVBoxLayout(self.stats_button)
            btn_layout.addWidget(stats_icon_widget)
            btn_layout.setAlignment(Qt.AlignCenter)
            btn_layout.setContentsMargins(0, 0, 0, 0)
        else:
            self.stats_button = QPushButton('📊')
        
        self.stats_button.setFixedSize(80, 80)
        self.stats_button.setEnabled(False)
        self.stats_button.setToolTip('View analysis statistics')
        self.stats_button.setCursor(Qt.PointingHandCursor)
        self.stats_button.clicked.connect(self.show_statistics_dialog)
        self.stats_button.setStyleSheet(f"""
            QPushButton {{
                background: #8B0000;
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 32px;
            }}
            QPushButton:hover {{
                background: #A52A2A;
                cursor: pointer;
            }}
            QPushButton:pressed {{
                background: #660000;
            }}
            QPushButton:disabled {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_muted']};
            }}
        """)
        stats_container.addWidget(self.stats_button)
        ar.addLayout(stats_container)
        
        ar.addStretch()
        actions_section.addLayout(ar)
        
        lay.addLayout(actions_section)
        
        lay.addSpacing(30)

        # Options - Centrado
        options_section = QVBoxLayout()
        options_section.setAlignment(Qt.AlignCenter)
        
        options_label_container = QHBoxLayout()
        options_label_container.addStretch()
        options_label_container.addWidget(section_label('Export Options'))
        options_label_container.addStretch()
        options_section.addLayout(options_label_container)
        options_section.addSpacing(8)

        opts_grid = QHBoxLayout()
        opts_grid.setSpacing(20)
        opts_grid.addStretch()

        self.keep_original_name = QCheckBox('Overwrite originals')
        self.disable_popups = QCheckBox('Silent export')
        self.optimize_pdf = QCheckBox('Optimize PDF')
        self.clean_pdf_links = QCheckBox('Clean existing links & JS')

        for cb in (self.keep_original_name, self.disable_popups, self.optimize_pdf, self.clean_pdf_links):
            opts_grid.addWidget(cb)

        opts_grid.addStretch()
        options_section.addLayout(opts_grid)
        
        lay.addLayout(options_section)
        lay.addStretch()

        # Hidden compat
        self.file_label = QLabel()
        self.scan_page_spinbox = QSpinBox()
        self.scan_page_spinbox.setRange(1, 999); self.scan_page_spinbox.setValue(2)
        self.scan_page_spinbox.setVisible(False)
        lay.addWidget(self.scan_page_spinbox)

        return w

    def _drop_zone_style(self, active):
        if active:
            return f"""
                QLabel {{
                    background: {COLORS['accent_dim']};
                    border: 2px dashed {COLORS['accent']};
                    border-radius: 14px;
                }}
            """
        return f"""
            QLabel {{
                background: {COLORS['bg_base']};
                border: 2px dashed {COLORS['border']};
                border-radius: 14px;
            }}
            QLabel:hover {{
                border-color: {COLORS['text_muted']};
            }}
        """

    # ─── PAGE: PATTERN ────────────────────────
    def _page_pattern(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 40, 0, 0)  # Añadir margen superior de 40px
        lay.setSpacing(0)

        # Contenedor centrado (sin centrado vertical)
        center_container = QWidget()
        center_layout = QVBoxLayout(center_container)
        center_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)  # Centrado horizontal, alineado arriba
        center_layout.setSpacing(32)

        # Título principal centrado
        title_container = QVBoxLayout()
        title_container.setAlignment(Qt.AlignCenter)
        title_container.setSpacing(8)
        
        main_title = QLabel('Reference Pattern')
        main_title.setAlignment(Qt.AlignCenter)
        main_title.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 24px;
            font-weight: 700;
            background: transparent;
        """)
        title_container.addWidget(main_title)
        
        subtitle = QLabel('Configure how references are detected in your PDFs')
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 14px;
            background: transparent;
        """)
        title_container.addWidget(subtitle)
        
        center_layout.addLayout(title_container)

        # Contenedor de controles centrado
        controls_container = QWidget()
        controls_container.setMaximumWidth(600)
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setSpacing(24)

        # Selector de estilo
        style_section = QVBoxLayout()
        style_section.setSpacing(12)
        
        style_label = QLabel('Style')
        style_label.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        """)
        style_section.addWidget(style_label)
        
        style_row = QHBoxLayout()
        style_row.setSpacing(16)
        
        self.pattern_combo = QComboBox()
        self.pattern_combo.addItems(list(self.REFERENCE_PATTERNS.keys()))
        self.pattern_combo.currentTextChanged.connect(self.on_pattern_changed)
        self.pattern_combo.setFixedHeight(28)  # Altura fija de 28px
        self.pattern_combo.setStyleSheet(f"""
            QComboBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 0px 16px;  /* Sin padding vertical */
                color: {COLORS['text_primary']};
                font-size: 12px;
                height: 28px;
                line-height: 28px;
            }}
            QComboBox:focus {{
                border-color: {COLORS['accent']};
            }}
        """)
        style_row.addWidget(self.pattern_combo, 1)
        
        self.pattern_example = QLabel('/1.0-A, /10.5-Z')
        self.pattern_example.setAlignment(Qt.AlignCenter)
        self.pattern_example.setStyleSheet(f"""
            background: {COLORS['accent_dim']};
            color: {COLORS['accent']};
            padding: 10px 16px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            font-family: 'Consolas', monospace;
        """)
        style_row.addWidget(self.pattern_example)
        
        help_btn = QPushButton('?')
        help_btn.setFixedSize(44, 44)
        help_btn.clicked.connect(self.show_pattern_help)
        help_btn.setStyleSheet(f"""
            QPushButton {{
                background: {COLORS['bg_elevated']};
                color: {COLORS['text_secondary']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                font-size: 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background: {COLORS['bg_hover']};
                border-color: {COLORS['accent']};
                color: {COLORS['accent']};
            }}
        """)
        style_row.addWidget(help_btn)
        
        style_section.addLayout(style_row)
        controls_layout.addLayout(style_section)

        # Patrón personalizado
        pattern_section = QVBoxLayout()
        pattern_section.setSpacing(12)
        
        pattern_label = QLabel('Custom Pattern')
        pattern_label.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        """)
        pattern_section.addWidget(pattern_label)
        
        self.custom_pattern_input = QLineEdit()
        self.custom_pattern_input.setPlaceholderText('e.g. /{P}.{C}-{F}  or  {P}-{F}.{C}')
        self.custom_pattern_input.setEnabled(False)
        self.custom_pattern_input.textChanged.connect(self.on_custom_pattern_changed)
        self.custom_pattern_input.setMinimumHeight(44)
        self.custom_pattern_input.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 14px;
                font-family: 'Consolas', monospace;
            }}
            QLineEdit:focus {{
                border-color: {COLORS['accent']};
            }}
            QLineEdit:disabled {{
                background: {COLORS['bg_surface']};
                color: {COLORS['text_muted']};
            }}
        """)
        pattern_section.addWidget(self.custom_pattern_input)
        
        # Hint y preview (ocultos inicialmente)
        self.pattern_hint = QLabel('💡  Placeholders: {P} = page  {C} = column  {F} = row')
        self.pattern_hint.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            background: {COLORS['bg_elevated']};
            padding: 8px 12px;
            border-radius: 6px;
        """)
        self.pattern_hint.hide()
        pattern_section.addWidget(self.pattern_hint)

        self.pattern_preview_label = QLabel()
        self.pattern_preview_label.setStyleSheet(f"""
            color: {COLORS['success']};
            font-size: 12px;
            font-family: 'Consolas', monospace;
            background: {COLORS['success_dim']};
            padding: 8px 12px;
            border-radius: 6px;
        """)
        self.pattern_preview_label.hide()
        pattern_section.addWidget(self.pattern_preview_label)
        
        controls_layout.addLayout(pattern_section)
        
        center_layout.addWidget(controls_container)
        
        # Añadir el contenedor al layout principal (sin stretch para que quede arriba)
        lay.addWidget(center_container)
        lay.addStretch()  # Solo stretch al final para empujar contenido hacia arriba
        
        return w

    # ─── PAGE: GRID ───────────────────────────
    def _page_grid(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Contenedor centrado
        center_container = QWidget()
        center_layout = QVBoxLayout(center_container)
        center_layout.setAlignment(Qt.AlignCenter)
        center_layout.setSpacing(40)

        # Título principal centrado
        title_container = QVBoxLayout()
        title_container.setAlignment(Qt.AlignCenter)
        title_container.setSpacing(8)
        
        main_title = QLabel('Grid Configuration')
        main_title.setAlignment(Qt.AlignCenter)
        main_title.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 24px;
            font-weight: 700;
            background: transparent;
        """)
        title_container.addWidget(main_title)
        
        subtitle = QLabel('Define the grid layout for reference detection')
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 14px;
            background: transparent;
        """)
        title_container.addWidget(subtitle)
        
        center_layout.addLayout(title_container)

        # Contenedor de controles centrado
        controls_container = QWidget()
        controls_container.setMaximumWidth(700)
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setSpacing(32)

        # Sección: Dimensiones
        dimensions_section = QVBoxLayout()
        dimensions_section.setSpacing(16)
        
        dimensions_label = QLabel('Dimensions')
        dimensions_label.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        """)
        dimensions_section.addWidget(dimensions_label)
        
        # Fila de columnas y filas
        dims_row = QHBoxLayout()
        dims_row.setSpacing(24)
        
        # Columnas
        cols_container = QVBoxLayout()
        cols_container.setSpacing(8)
        cols_label = QLabel('Columns')
        cols_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        cols_container.addWidget(cols_label)
        
        self.cols_spinbox = QSpinBox()
        self.cols_spinbox.setRange(1, 50)
        self.cols_spinbox.setValue(10)
        self.cols_spinbox.valueChanged.connect(self.update_size_placeholders)
        self.cols_spinbox.setMinimumHeight(44)
        self.cols_spinbox.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 16px;
                font-weight: 600;
            }}
            QSpinBox:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        cols_container.addWidget(self.cols_spinbox)
        dims_row.addLayout(cols_container)
        
        # Filas
        rows_container = QVBoxLayout()
        rows_container.setSpacing(8)
        rows_label = QLabel('Rows')
        rows_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        rows_container.addWidget(rows_label)
        
        rows_input_container = QHBoxLayout()
        rows_input_container.setSpacing(12)
        
        self.rows_spinbox = QSpinBox()
        self.rows_spinbox.setRange(1, 26)
        self.rows_spinbox.setValue(8)
        self.rows_spinbox.valueChanged.connect(self.update_rows_info)
        self.rows_spinbox.valueChanged.connect(self.update_size_placeholders)
        self.rows_spinbox.setMinimumHeight(44)
        self.rows_spinbox.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 16px;
                font-weight: 600;
            }}
            QSpinBox:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        rows_input_container.addWidget(self.rows_spinbox)
        
        self.rows_info_label = QLabel('A–H')
        self.rows_info_label.setAlignment(Qt.AlignCenter)
        self.rows_info_label.setStyleSheet(f"""
            background: {COLORS['purple_dim']};
            color: {COLORS['purple']};
            padding: 10px 16px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            font-family: 'Consolas', monospace;
        """)
        rows_input_container.addWidget(self.rows_info_label)
        
        rows_container.addLayout(rows_input_container)
        dims_row.addLayout(rows_container)
        
        dimensions_section.addLayout(dims_row)
        controls_layout.addLayout(dimensions_section)

        # Sección: Márgenes
        margins_section = QVBoxLayout()
        margins_section.setSpacing(16)
        
        margins_label = QLabel('Margins')
        margins_label.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        """)
        margins_section.addWidget(margins_label)
        
        # Fila de márgenes
        margins_row = QHBoxLayout()
        margins_row.setSpacing(24)
        
        # Margen izquierdo
        margin_left_container = QVBoxLayout()
        margin_left_container.setSpacing(8)
        margin_left_label = QLabel('Left')
        margin_left_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        margin_left_container.addWidget(margin_left_label)
        
        self.margin_left_spinbox = QSpinBox()
        self.margin_left_spinbox.setRange(0, 30)
        self.margin_left_spinbox.setValue(5)
        self.margin_left_spinbox.setSuffix('%')
        self.margin_left_spinbox.setMinimumHeight(44)
        self.margin_left_spinbox.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 14px;
            }}
            QSpinBox:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        margin_left_container.addWidget(self.margin_left_spinbox)
        margins_row.addLayout(margin_left_container)
        
        # Margen superior
        margin_top_container = QVBoxLayout()
        margin_top_container.setSpacing(8)
        margin_top_label = QLabel('Top')
        margin_top_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        margin_top_container.addWidget(margin_top_label)
        
        self.margin_top_spinbox = QSpinBox()
        self.margin_top_spinbox.setRange(0, 30)
        self.margin_top_spinbox.setValue(5)
        self.margin_top_spinbox.setSuffix('%')
        self.margin_top_spinbox.setMinimumHeight(44)
        self.margin_top_spinbox.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 14px;
            }}
            QSpinBox:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        margin_top_container.addWidget(self.margin_top_spinbox)
        margins_row.addLayout(margin_top_container)
        
        margins_section.addLayout(margins_row)
        controls_layout.addLayout(margins_section)

        # Sección: Tamaños personalizados
        sizes_section = QVBoxLayout()
        sizes_section.setSpacing(16)
        
        sizes_label = QLabel('Custom Sizes (Optional)')
        sizes_label.setStyleSheet(f"""
            color: {COLORS['text_muted']};
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
        """)
        sizes_section.addWidget(sizes_label)
        
        # Anchos de columnas
        col_widths_container = QVBoxLayout()
        col_widths_container.setSpacing(8)
        col_widths_label = QLabel('Column Widths')
        col_widths_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        col_widths_container.addWidget(col_widths_label)
        
        self.col_sizes_input = QLineEdit()
        self.col_sizes_input.setPlaceholderText('1,1,1,… (relative values, leave empty for equal)')
        self.col_sizes_input.setMinimumHeight(44)
        self.col_sizes_input.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 14px;
                font-family: 'Consolas', monospace;
            }}
            QLineEdit:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        col_widths_container.addWidget(self.col_sizes_input)
        sizes_section.addLayout(col_widths_container)
        
        # Alturas de filas
        row_heights_container = QVBoxLayout()
        row_heights_container.setSpacing(8)
        row_heights_label = QLabel('Row Heights')
        row_heights_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        row_heights_container.addWidget(row_heights_label)
        
        self.row_sizes_input = QLineEdit()
        self.row_sizes_input.setPlaceholderText('1,1,1,… (relative values, leave empty for equal)')
        self.row_sizes_input.setMinimumHeight(44)
        self.row_sizes_input.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 16px;
                color: {COLORS['text_primary']};
                font-size: 14px;
                font-family: 'Consolas', monospace;
            }}
            QLineEdit:focus {{
                border-color: {COLORS['purple']};
            }}
        """)
        row_heights_container.addWidget(self.row_sizes_input)
        sizes_section.addLayout(row_heights_container)
        
        controls_layout.addLayout(sizes_section)
        
        center_layout.addWidget(controls_container)
        
        # Añadir el contenedor centrado al layout principal
        lay.addStretch()
        lay.addWidget(center_container)
        lay.addStretch()
        
        return w

    # ─── PAGE: STYLE ──────────────────────────
    def _page_style(self):
        w = QWidget()
        main_layout = QVBoxLayout(w)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Crear scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet(f"""
            QScrollArea {{
                border: none;
                background: {COLORS['bg_base']};
            }}
        """)
        
        # Contenedor de contenido scrollable
        scroll_content = QWidget()
        lay = QVBoxLayout(scroll_content)
        lay.setContentsMargins(0, 20, 0, 20)
        lay.setSpacing(0)

        # Contenedor centrado
        center_container = QWidget()
        center_layout = QVBoxLayout(center_container)
        center_layout.setAlignment(Qt.AlignCenter)
        center_layout.setSpacing(40)

        # Título principal centrado
        title_container = QVBoxLayout()
        title_container.setAlignment(Qt.AlignCenter)
        title_container.setSpacing(8)
        
        main_title = QLabel('Highlight Style')
        main_title.setAlignment(Qt.AlignCenter)
        main_title.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 24px;
            font-weight: 700;
            background: transparent;
        """)
        title_container.addWidget(main_title)
        
        subtitle = QLabel('Customize the appearance of reference highlights')
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 14px;
            background: transparent;
        """)
        title_container.addWidget(subtitle)
        
        center_layout.addLayout(title_container)

        # Contenedor de controles centrado
        controls_container = QWidget()
        controls_container.setMaximumWidth(1200)  # Aumentar a 1200px para mejor distribución
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setSpacing(20)  # Reducir espacio entre secciones

        # Sección: Apariencia básica
        basic_section = QVBoxLayout()
        basic_section.setSpacing(12)
        
        basic_label = QLabel('Basic Appearance')
        basic_label.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 14px;
            font-weight: 700;
        """)
        basic_section.addWidget(basic_label)
        
        basic_desc = QLabel('Configure the border color, width, line style, and blink speed')
        basic_desc.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 12px;
            margin-bottom: 8px;
        """)
        basic_section.addWidget(basic_desc)
        
        # Grid layout para organización perfecta
        basic_grid = QGridLayout()
        basic_grid.setSpacing(30)
        basic_grid.setVerticalSpacing(8)  # Espacio vertical reducido entre label y control
        basic_grid.setColumnStretch(0, 1)
        basic_grid.setColumnStretch(1, 1)
        basic_grid.setColumnStretch(2, 1)
        
        # Fila 1: Labels
        color_label = QLabel('Color')
        color_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        basic_grid.addWidget(color_label, 0, 0, Qt.AlignLeft)
        
        width_label = QLabel('Width')
        width_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        basic_grid.addWidget(width_label, 0, 1, Qt.AlignLeft)
        
        line_label = QLabel('Line Style')
        line_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        basic_grid.addWidget(line_label, 0, 2, Qt.AlignLeft)
        
        # Fila 2: Controles
        self.color_combo = self._make_combo('color_combo', ['Red','Green','Blue','Yellow','Orange','Magenta','Cyan'], 'Red')
        self.color_combo.setFixedWidth(200)
        basic_grid.addWidget(self.color_combo, 1, 0)
        
        self.line_width_spinbox = self._make_spinbox('line_width_spinbox', 1, 10, 3, ' px')
        self.line_width_spinbox.setFixedWidth(200)
        basic_grid.addWidget(self.line_width_spinbox, 1, 1)
        
        self.line_style_combo = self._make_combo('line_style_combo', ['Solid','Dashed','Dotted'])
        self.line_style_combo.setFixedWidth(200)
        basic_grid.addWidget(self.line_style_combo, 1, 2)
        
        # Fila 3: Espacio
        basic_grid.setRowMinimumHeight(2, 20)
        
        # Fila 4: Blink Speed Label (centrado)
        blink_label = QLabel('Blink Speed')
        blink_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        blink_label.setAlignment(Qt.AlignCenter)
        basic_grid.addWidget(blink_label, 3, 1, Qt.AlignCenter)
        
        # Fila 5: Blink Speed Control (centrado)
        self.blink_speed_combo = self._make_combo('blink_speed_combo', ['Fast','Normal','Slow','None'], 'Normal')
        self.blink_speed_combo.setFixedWidth(200)
        basic_grid.addWidget(self.blink_speed_combo, 4, 1)
        
        basic_section.addLayout(basic_grid)
        controls_layout.addLayout(basic_section)
        
        # Línea separadora naranja
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setStyleSheet(f"background: {COLORS['accent']}; max-height: 2px; border: none; margin: 10px 0;")
        controls_layout.addWidget(separator1)

        # Sección: Animación y efectos
        animation_section = QVBoxLayout()
        animation_section.setSpacing(12)
        
        animation_label = QLabel('Animation & Effects')
        animation_label.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 14px;
            font-weight: 700;
        """)
        animation_section.addWidget(animation_label)
        
        animation_desc = QLabel('Set the highlight duration, fill style, animation type, and opacity')
        animation_desc.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 12px;
            margin-bottom: 8px;
        """)
        animation_section.addWidget(animation_desc)
        
        # Grid layout para organización perfecta
        anim_grid = QGridLayout()
        anim_grid.setSpacing(30)
        anim_grid.setVerticalSpacing(8)  # Espacio vertical reducido entre label y control
        anim_grid.setColumnStretch(0, 1)
        anim_grid.setColumnStretch(1, 1)
        anim_grid.setColumnStretch(2, 1)
        
        # Fila 1: Labels
        duration_label = QLabel('Duration')
        duration_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        anim_grid.addWidget(duration_label, 0, 0, Qt.AlignLeft)
        
        fill_label = QLabel('Fill')
        fill_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        anim_grid.addWidget(fill_label, 0, 1, Qt.AlignLeft)
        
        anim_type_label = QLabel('Animation')
        anim_type_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        anim_grid.addWidget(anim_type_label, 0, 2, Qt.AlignLeft)
        
        # Fila 2: Controles
        # Duration spinbox con decimales para permitir duraciones más cortas
        from PyQt5.QtWidgets import QDoubleSpinBox
        self.duration_spinbox = QDoubleSpinBox()
        self.duration_spinbox.setRange(0.1, 30)  # Mínimo 0.1 segundos
        self.duration_spinbox.setValue(5)
        self.duration_spinbox.setSingleStep(0.5)  # Incrementos de 0.5 segundos
        self.duration_spinbox.setDecimals(1)  # Un decimal
        self.duration_spinbox.setSuffix(' s')
        self.duration_spinbox.setMinimumHeight(44)
        self.duration_spinbox.setStyleSheet(f"""
            QDoubleSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 13px;
                color: {COLORS['text_primary']};
            }}
            QDoubleSpinBox:hover {{
                border-color: {COLORS['accent']};
            }}
            QDoubleSpinBox:focus {{
                border-color: {COLORS['accent']};
                outline: none;
            }}
            QDoubleSpinBox::up-button, QDoubleSpinBox::down-button {{
                width: 20px;
                border: none;
                background: transparent;
            }}
            QDoubleSpinBox::up-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-bottom: 4px solid {COLORS['text_secondary']};
            }}
            QDoubleSpinBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid {COLORS['text_secondary']};
            }}
        """)
        self.duration_spinbox.setFixedWidth(200)
        anim_grid.addWidget(self.duration_spinbox, 1, 0)
        
        self.fill_combo = self._make_combo('fill_combo', ['None','Semi-transparent','Solid'])
        self.fill_combo.setFixedWidth(200)
        anim_grid.addWidget(self.fill_combo, 1, 1)
        
        self.animation_combo = self._make_combo('animation_combo', ['Blink','Fade','Pulse','None'], 'Blink')
        self.animation_combo.setFixedWidth(200)
        anim_grid.addWidget(self.animation_combo, 1, 2)
        
        # Fila 3: Espacio
        anim_grid.setRowMinimumHeight(2, 20)
        
        # Fila 4: Opacity Label (centrado)
        opacity_label = QLabel('Opacity')
        opacity_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        opacity_label.setAlignment(Qt.AlignCenter)
        anim_grid.addWidget(opacity_label, 3, 1, Qt.AlignCenter)
        
        # Fila 5: Opacity Control (centrado)
        self.opacity_spinbox = self._make_spinbox('opacity_spinbox', 10, 100, 100, '%')
        self.opacity_spinbox.setFixedWidth(200)
        anim_grid.addWidget(self.opacity_spinbox, 4, 1)
        
        animation_section.addLayout(anim_grid)
        controls_layout.addLayout(animation_section)
        
        # Línea separadora naranja
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.HLine)
        separator2.setStyleSheet(f"background: {COLORS['accent']}; max-height: 2px; border: none; margin: 10px 0;")
        controls_layout.addWidget(separator2)

        # Sección: Avanzado
        advanced_section = QVBoxLayout()
        advanced_section.setSpacing(12)
        
        advanced_label = QLabel('Advanced Options')
        advanced_label.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 14px;
            font-weight: 700;
        """)
        advanced_section.addWidget(advanced_label)
        
        advanced_desc = QLabel('Fine-tune fill color, corner radius, margin, and visual effects')
        advanced_desc.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 12px;
            margin-bottom: 8px;
        """)
        advanced_section.addWidget(advanced_desc)
        
        # Grid layout para 4 columnas perfectamente organizadas
        adv_grid = QGridLayout()
        adv_grid.setSpacing(25)
        adv_grid.setVerticalSpacing(8)  # Espacio vertical reducido entre label y control
        adv_grid.setColumnStretch(0, 1)
        adv_grid.setColumnStretch(1, 1)
        adv_grid.setColumnStretch(2, 1)
        adv_grid.setColumnStretch(3, 1)
        
        # Fila 1: Labels
        fill_color_label = QLabel('Fill Color')
        fill_color_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        adv_grid.addWidget(fill_color_label, 0, 0, Qt.AlignLeft)
        
        radius_label = QLabel('Corner Radius')
        radius_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        adv_grid.addWidget(radius_label, 0, 1, Qt.AlignLeft)
        
        margin_label = QLabel('Margin')
        margin_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        adv_grid.addWidget(margin_label, 0, 2, Qt.AlignLeft)
        
        effect_label = QLabel('Effect')
        effect_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 13px; font-weight: 500;")
        adv_grid.addWidget(effect_label, 0, 3, Qt.AlignLeft)
        
        # Fila 2: Controles
        self.fill_color_combo = self._make_combo('fill_color_combo', ['Same as border','Red','Green','Blue','Yellow','Orange','Magenta','Cyan','White','Black'])
        self.fill_color_combo.setFixedWidth(180)
        adv_grid.addWidget(self.fill_color_combo, 1, 0)
        
        self.corner_radius_spinbox = self._make_spinbox('corner_radius_spinbox', 0, 20, 0, ' px')
        self.corner_radius_spinbox.setFixedWidth(180)
        adv_grid.addWidget(self.corner_radius_spinbox, 1, 1)
        
        self.rect_margin_spinbox = self._make_spinbox('rect_margin_spinbox', -20, 20, 0, ' px')
        self.rect_margin_spinbox.setFixedWidth(180)
        adv_grid.addWidget(self.rect_margin_spinbox, 1, 2)
        
        self.effect_combo = self._make_combo('effect_combo', ['None','Soft shadow','Glow'])
        self.effect_combo.setFixedWidth(180)
        adv_grid.addWidget(self.effect_combo, 1, 3)
        
        advanced_section.addLayout(adv_grid)
        controls_layout.addLayout(advanced_section)
        
        # Línea separadora naranja
        separator3 = QFrame()
        separator3.setFrameShape(QFrame.HLine)
        separator3.setStyleSheet(f"background: {COLORS['accent']}; max-height: 2px; border: none; margin: 10px 0;")
        controls_layout.addWidget(separator3)

        # Sección: Vista previa
        preview_section = QVBoxLayout()
        preview_section.setSpacing(12)
        
        preview_label = QLabel('Preview')
        preview_label.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            font-size: 14px;
            font-weight: 700;
        """)
        preview_section.addWidget(preview_label)
        
        preview_desc = QLabel('See how your highlight style will look in the PDF')
        preview_desc.setStyleSheet(f"""
            color: {COLORS['text_secondary']};
            font-size: 12px;
            margin-bottom: 8px;
        """)
        preview_section.addWidget(preview_desc)
        
        # Contenedor de preview centrado
        preview_container = QWidget()
        preview_container.setStyleSheet(f"""
            background: {COLORS['bg_base']};
            border-radius: 12px;
            padding: 20px;
        """)
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setAlignment(Qt.AlignCenter)
        
        self.style_preview = QLabel('━━━━━━━━━━━━━━━━━━')
        self.style_preview.setAlignment(Qt.AlignCenter)
        self.style_preview.setStyleSheet(f"""
            color: {COLORS['danger']};
            font-size: 24px;
            font-weight: 700;
            letter-spacing: 3px;
            background: transparent;
        """)
        preview_layout.addWidget(self.style_preview)
        
        preview_section.addWidget(preview_container)
        controls_layout.addLayout(preview_section)
        
        center_layout.addWidget(controls_container)
        
        # Añadir el contenedor centrado al layout principal
        lay.addStretch()
        lay.addWidget(center_container)
        lay.addStretch()
        
        # Añadir el contenido scrollable al scroll area
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)

        # Connect signals
        self.color_combo.currentTextChanged.connect(self.update_style_preview)
        for attr in ('line_width_spinbox','line_style_combo','blink_speed_combo',
                     'duration_spinbox','fill_combo','animation_combo','opacity_spinbox',
                     'fill_color_combo','corner_radius_spinbox','rect_margin_spinbox','effect_combo'):
            w2 = getattr(self, attr)
            if hasattr(w2, 'currentTextChanged'):
                w2.currentTextChanged.connect(self.update_style_preview)
                w2.currentTextChanged.connect(self.save_styles_config)
            elif hasattr(w2, 'valueChanged'):
                w2.valueChanged.connect(self.update_style_preview)
                w2.valueChanged.connect(self.save_styles_config)

        return w

    def _make_combo(self, attr, items, default=None):
        cb = QComboBox()
        cb.addItems(items)
        if default and default in items:
            cb.setCurrentText(default)
        cb.setFixedHeight(28)  # Altura fija de 28px
        cb.setStyleSheet(f"""
            QComboBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 0px 12px;  /* Sin padding vertical */
                font-size: 12px;
                color: {COLORS['text_primary']};
                height: 28px;
                line-height: 28px;
            }}
            QComboBox:hover {{
                border-color: {COLORS['accent']};
            }}
            QComboBox:focus {{
                border-color: {COLORS['accent']};
                outline: none;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 30px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid {COLORS['text_secondary']};
                margin-right: 8px;
            }}
        """)
        setattr(self, attr, cb)
        return cb

    def _make_spinbox(self, attr, lo, hi, val, suffix=''):
        sb = QSpinBox()
        sb.setRange(lo, hi); sb.setValue(val)
        if suffix: sb.setSuffix(suffix)
        sb.setMinimumHeight(44)
        sb.setStyleSheet(f"""
            QSpinBox {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 13px;
                color: {COLORS['text_primary']};
            }}
            QSpinBox:hover {{
                border-color: {COLORS['accent']};
            }}
            QSpinBox:focus {{
                border-color: {COLORS['accent']};
                outline: none;
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                width: 20px;
                border: none;
                background: transparent;
            }}
            QSpinBox::up-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-bottom: 4px solid {COLORS['text_secondary']};
            }}
            QSpinBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid {COLORS['text_secondary']};
            }}
        """)
        setattr(self, attr, sb)
        return sb

    # ─── DRAG & DROP ──────────────────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            if any(u.toLocalFile().lower().endswith('.pdf') for u in event.mimeData().urls()):
                event.acceptProposedAction()
                self.drop_zone_container.setStyleSheet(self._drop_zone_style(True))
            else:
                event.ignore()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.drop_zone_container.setStyleSheet(self._drop_zone_style(False))

    def dropEvent(self, event):
        pdfs = [u.toLocalFile() for u in event.mimeData().urls()
                if u.toLocalFile().lower().endswith('.pdf')]
        if pdfs: self.add_pdf_files(pdfs)
        self.drop_zone_container.setStyleSheet(self._drop_zone_style(False))

    def add_pdf_files(self, file_paths):
        added = 0
        for fp in file_paths:
            if fp not in self.pdf_paths:
                self.pdf_paths.append(fp)
                # Obtener nombre del archivo
                filename = os.path.basename(fp)
                
                # Crear widget personalizado para el item
                item_widget = QWidget()
                item_widget.setFixedHeight(40)
                item_layout = QHBoxLayout(item_widget)
                item_layout.setContentsMargins(8, 10, 8, 10)
                item_layout.setSpacing(8)
                
                # Añadir icono SVG
                if 'pdf_icon' in SVG_ICONS:
                    pdf_icon = create_svg_icon(SVG_ICONS['pdf_icon'], 24)  # Aumentado de 18 a 24px
                    pdf_icon.setFixedSize(24, 24)
                    pdf_icon.setAlignment(Qt.AlignCenter)
                    item_layout.addWidget(pdf_icon)
                
                # Crear label temporal para medir el texto
                temp_label = QLabel(filename)
                temp_label.setStyleSheet(f"""
                    QLabel {{
                        font-size: 13px;
                        font-weight: 500;
                    }}
                """)
                
                # Calcular ancho disponible (ancho de lista - icono - márgenes - spacing)
                available_width = 500 - 24 - 16 - 8  # 452px disponibles para texto (icono ahora es 24px)
                text_width = temp_label.fontMetrics().boundingRect(filename).width()
                
                # Solo truncar si el texto no cabe
                display_filename = filename
                if text_width > available_width:
                    # Calcular cuántos caracteres caben
                    chars_that_fit = len(filename)
                    while temp_label.fontMetrics().boundingRect(display_filename + "...").width() > available_width and chars_that_fit > 10:
                        chars_that_fit -= 1
                        display_filename = filename[:chars_that_fit] + "..."
                
                # Añadir texto del archivo
                file_label = QLabel(display_filename)
                file_label.setStyleSheet(f"""
                    QLabel {{
                        color: {COLORS['text_primary']};
                        font-size: 13px;
                        font-weight: 500;
                        background: transparent;
                        border: none;
                        padding: 0px;
                        margin: 0px;
                    }}
                """)
                file_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                file_label.setWordWrap(False)
                file_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                file_label.setToolTip(filename)  # Mostrar nombre completo en tooltip
                item_layout.addWidget(file_label, 1)
                
                # Crear el item de la lista
                item = QListWidgetItem()
                item.setData(Qt.UserRole, fp)
                item.setSizeHint(QSize(-1, 40))  # Altura aumentada a 40px
                
                # Añadir a la lista
                self.pdf_list.addItem(item)
                self.pdf_list.setItemWidget(item, item_widget)
                added += 1
        if added:
            # Cambiar contenido del drop zone: ocultar contenido inicial y mostrar lista
            self.drop_zone_content.hide()
            self.drop_zone_file_list.show()
            
            self._update_badges()
            self.detect_button.setEnabled(True)
            self.stats_button.setEnabled(False)  # Inicialmente deshabilitado hasta que haya análisis
            self.visual_editor_button.setEnabled(True)
            if self.pdf_paths: self.pdf_path = self.pdf_paths[0]
            self.statusBar().showMessage(f'{added} file(s) added  ·  Total: {len(self.pdf_paths)}')
            self.load_saved_grid_config()

    def _update_badges(self):
        n = len(self.pdf_paths)
        self.pdf_count_badge.setText(f'{n} file{"s" if n != 1 else ""}')
        self.pdf_count_badge.setStyleSheet(
            self.pdf_count_badge.styleSheet().replace(COLORS['bg_elevated'], COLORS['accent_dim'])
            if n > 0 else self.pdf_count_badge.styleSheet()
        )

    def update_pdf_count(self):
        self._update_badges()

    def clear_pdf_list(self):
        self.pdf_paths.clear(); self.pdf_list.clear()
        self.all_references.clear(); self.references.clear()
        self.pdf_path = None
        
        # Volver a mostrar el contenido del drop zone
        self.drop_zone_file_list.hide()
        self.drop_zone_content.show()
        
        self._update_badges()
        self.ref_badge.setText('0 refs')
        self.detect_button.setEnabled(False)
        self.stats_button.setEnabled(False)
        self.generate_button.setEnabled(False)
        self.table.setRowCount(0)
        self.statusBar().showMessage('File list cleared')

    def remove_selected_pdfs(self):
        for item in self.pdf_list.selectedItems():
            fp = item.data(Qt.UserRole)
            if fp in self.pdf_paths: self.pdf_paths.remove(fp)
            if fp in self.all_references: del self.all_references[fp]
            self.pdf_list.takeItem(self.pdf_list.row(item))
        self._update_badges()
        if not self.pdf_paths:
            # Volver a mostrar el contenido del drop zone si no hay archivos
            self.drop_zone_file_list.hide()
            self.drop_zone_content.show()
            
            self.detect_button.setEnabled(False)
            self.stats_button.setEnabled(False)
            self.generate_button.setEnabled(False)
            self.pdf_path = None
        else:
            self.pdf_path = self.pdf_paths[0]

    def select_pdf(self):
        fps, _ = QFileDialog.getOpenFileNames(self, 'Select PDF Files', '', 'PDF Files (*.pdf)')
        if fps: self.add_pdf_files(fps)

    # ─── GRID CONFIG ──────────────────────────
    def show_grid_template_menu(self):
        """Mostrar menú para seleccionar plantillas de grid"""
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background: {COLORS['bg_elevated']};
                border: 2px solid {COLORS['accent']};
                border-radius: 10px;
                padding: 8px;
            }}
            QMenu::item {{
                color: {COLORS['text_primary']};
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 13px;
            }}
            QMenu::item:selected {{
                background: {COLORS['accent_dim']};
                color: {COLORS['accent']};
            }}
            QMenu::separator {{
                height: 1px;
                background: {COLORS['border']};
                margin: 6px 10px;
            }}
        """)
        
        # Cargar plantillas guardadas
        templates = self.load_grid_templates()
        
        # Añadir plantillas existentes
        if templates:
            for name, template in templates.items():
                action = menu.addAction(f"● {name} ({template['cols']}×{template['rows']})")
                action.triggered.connect(lambda checked, n=name: self.load_grid_template(n))
            menu.addSeparator()
            
            # Opción de gestión
            manage_action = menu.addAction("⚙️ Manage Templates...")
            manage_action.triggered.connect(self.manage_grid_templates)
        else:
            no_templates = menu.addAction("No templates saved")
            no_templates.setEnabled(False)
        
        menu.exec_(self.grid_template_btn.mapToGlobal(self.grid_template_btn.rect().topLeft()))
    
    def load_grid_templates(self):
        """Cargar plantillas de grid guardadas"""
        templates_file = os.path.join(get_app_path(), 'grid_templates.json')
        if os.path.exists(templates_file):
            try:
                with open(templates_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f'Error loading templates: {e}')
        return {}
    
    def save_grid_templates(self, templates):
        """Guardar plantillas de grid"""
        templates_file = os.path.join(get_app_path(), 'grid_templates.json')
        try:
            with open(templates_file, 'w') as f:
                json.dump(templates, f, indent=2)
        except Exception as e:
            print(f'Error saving templates: {e}')
    
    def load_grid_template(self, name):
        """Cargar una plantilla de grid"""
        templates = self.load_grid_templates()
        if name not in templates:
            self.show_warning_dialog('Template Not Found', f'Template "{name}" not found.')
            return
        
        template = templates[name]
        
        # Aplicar configuración
        self.column_positions = template['column_positions']
        self.row_positions = template['row_positions']
        self.grid_detected = True
        
        # Actualizar controles
        self.cols_spinbox.setValue(template['cols'])
        self.rows_spinbox.setValue(template['rows'])
        
        if template.get('col_sizes'):
            self.col_sizes_input.setText(template['col_sizes'])
        if template.get('row_sizes'):
            self.row_sizes_input.setText(template['row_sizes'])
        
        # Actualizar márgenes si hay información de página
        pw = template.get('page_width', 0)
        ph = template.get('page_height', 0)
        if pw and self.column_positions:
            self.margin_left_spinbox.setValue(max(0, min(int(self.column_positions[0]/pw*100), 30)))
        if ph and self.row_positions:
            self.margin_top_spinbox.setValue(max(0, min(int(self.row_positions[0]/ph*100), 30)))
        
        # Guardar como configuración actual
        self.save_grid_config()
        
        # Actualizar botón
        self.grid_template_btn.setText(f'● {name} ({template["cols"]}×{template["rows"]})')
        self.grid_template_btn.setStyleSheet(f"""
            QPushButton {{
                color: {COLORS['accent']};
                font-size: 10px;
                font-weight: 700;
                padding: 8px 6px;
                background: {COLORS['accent_dim']};
                border: 1px solid {COLORS['accent']};
                border-radius: 8px;
                text-align: center;
            }}
            QPushButton:hover {{
                background: {COLORS['accent_hover']};
                color: {COLORS['bg_base']};
            }}
        """)
        
        self.statusBar().showMessage(f'Grid template "{name}" loaded')
    
    def manage_grid_templates(self):
        """Diálogo para gestionar plantillas"""
        templates = self.load_grid_templates()
        if not templates:
            return
        
        d = QDialog(self)
        d.setWindowTitle('Manage Grid Templates')
        d.setMinimumSize(500, 400)
        d.setStyleSheet(f"QDialog {{ background: {COLORS['bg_base']}; }}")
        
        lay = QVBoxLayout(d)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(16)
        
        title = QLabel('Grid Templates')
        title.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 18px; font-weight: 700;")
        lay.addWidget(title)
        
        # Lista de plantillas
        template_list = QListWidget()
        template_list.setStyleSheet(f"""
            QListWidget {{
                background: {COLORS['bg_surface']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                padding: 8px;
            }}
            QListWidget::item {{
                padding: 12px;
                border-radius: 6px;
                margin: 2px;
            }}
            QListWidget::item:selected {{
                background: {COLORS['accent_dim']};
                color: {COLORS['accent']};
            }}
        """)
        
        for name, template in templates.items():
            template_list.addItem(f"{name} — {template['cols']}×{template['rows']} grid")
        
        lay.addWidget(template_list)
        
        # Botones
        btn_layout = QHBoxLayout()
        
        delete_btn = make_btn('Delete Selected', 'danger', 'sm')
        delete_btn.clicked.connect(lambda: self.delete_template(template_list, d))
        btn_layout.addWidget(delete_btn)
        
        btn_layout.addStretch()
        
        close_btn = make_btn('Close', 'accent', 'sm')
        close_btn.clicked.connect(d.accept)
        btn_layout.addWidget(close_btn)
        
        lay.addLayout(btn_layout)
        d.exec_()
    
    def delete_template(self, list_widget, dialog):
        """Eliminar plantilla seleccionada"""
        current_item = list_widget.currentItem()
        if not current_item:
            return
        
        # Extraer nombre de la plantilla
        text = current_item.text()
        name = text.split(' — ')[0]
        
        reply = QMessageBox.question(dialog, 'Delete Template', 
                                    f'Delete template "{name}"?',
                                    QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        
        # Eliminar plantilla
        templates = self.load_grid_templates()
        if name in templates:
            del templates[name]
            self.save_grid_templates(templates)
            list_widget.takeItem(list_widget.row(current_item))
            self.statusBar().showMessage(f'Template "{name}" deleted')
    
    def save_grid_config(self):
        """Guardar configuración de grid actual"""
        cp = os.path.join(get_app_path(), 'grid_config.json')
        try:
            config = {
                'column_lines': self.column_positions,
                'row_lines': self.row_positions,
                'page_width': getattr(self, 'grid_page_width', 0),
                'page_height': getattr(self, 'grid_page_height', 0)
            }
            with open(cp, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f'Error saving grid config: {e}')
    
    def load_saved_grid_config(self):
        cp = os.path.join(get_app_path(), 'grid_config.json')
        if not os.path.exists(cp):
            self.grid_detected = False; self.column_positions = []; self.row_positions = []
            self.grid_template_btn.setText('○  No grid configured')
            self.grid_template_btn.setStyleSheet(f"""
                QPushButton {{
                    color: {COLORS['text_muted']};
                    font-size: 10px;
                    padding: 8px 6px;
                    background: {COLORS['bg_elevated']};
                    border: 1px solid {COLORS['border']};
                    border-radius: 8px;
                    text-align: center;
                }}
                QPushButton:hover {{
                    background: {COLORS['bg_hover']};
                    border-color: {COLORS['accent']};
                }}
            """)
            return
        try:
            with open(cp, 'r') as f: d = json.load(f)
            self.column_positions = d.get('column_lines', [])
            self.row_positions = d.get('row_lines', [])
            self.grid_detected = len(self.column_positions) > 1 and len(self.row_positions) > 1
            if self.grid_detected:
                nc, nr = len(self.column_positions) - 1, len(self.row_positions) - 1
                self.cols_spinbox.setValue(nc); self.rows_spinbox.setValue(nr)
                pw, ph = d.get('page_width', 0), d.get('page_height', 0)
                if pw and self.column_positions:
                    self.margin_left_spinbox.setValue(max(0, min(int(self.column_positions[0]/pw*100), 30)))
                if ph and self.row_positions:
                    self.margin_top_spinbox.setValue(max(0, min(int(self.row_positions[0]/ph*100), 30)))
                self._compute_relative_sizes()
                self.grid_template_btn.setText(f'● Grid {nc}×{nr}')
                self.grid_template_btn.setStyleSheet(f"""
                    QPushButton {{
                        color: {COLORS['accent']};
                        font-size: 11px;
                        font-weight: 700;
                        padding: 8px 6px;
                        background: {COLORS['accent_dim']};
                        border: 1px solid {COLORS['accent']};
                        border-radius: 8px;
                        text-align: center;
                    }}
                    QPushButton:hover {{
                        background: {COLORS['accent_hover']};
                        color: {COLORS['bg_base']};
                    }}
                """)
        except Exception as e:
            print(f'Grid config load error: {e}')
            self.grid_template_btn.setText('○  No grid configured')
        self.load_styles_config()

    def _compute_relative_sizes(self):
        if len(self.column_positions) > 1:
            ws = [self.column_positions[i+1] - self.column_positions[i] for i in range(len(self.column_positions)-1)]
            mn = min(ws)
            if mn > 0: self.col_sizes_input.setText(','.join(str(round(w/mn, 2)) for w in ws))
        if len(self.row_positions) > 1:
            hs = [self.row_positions[i+1] - self.row_positions[i] for i in range(len(self.row_positions)-1)]
            mn = min(hs)
            if mn > 0: self.row_sizes_input.setText(','.join(str(round(h/mn, 2)) for h in hs))

    def load_styles_config(self):
        cp = os.path.join(get_app_path(), 'styles_config.json')
        if not os.path.exists(cp): return
        try:
            with open(cp) as f: c = json.load(f)
            for attr, key in [
                ('pattern_combo','pattern'), ('color_combo','rect_color'),
                ('line_style_combo','line_style'), ('blink_speed_combo','blink_speed'),
                ('fill_combo','fill_style'), ('animation_combo','animation_type'),
                ('fill_color_combo','fill_color'), ('effect_combo','effect')]:
                w = getattr(self, attr, None)
                if w and key in c:
                    i = w.findText(c[key])
                    if i >= 0: w.setCurrentIndex(i)
            for attr, key in [
                ('line_width_spinbox','line_width'), ('duration_spinbox','duration'),
                ('opacity_spinbox','opacity'), ('corner_radius_spinbox','corner_radius'),
                ('rect_margin_spinbox','rect_margin')]:
                w = getattr(self, attr, None)
                if w and key in c: w.setValue(c[key])
            if 'custom_pattern' in c: self.custom_pattern_input.setText(c['custom_pattern'])
            for attr, key in [('keep_original_name','keep_original_name'),
                               ('disable_popups','disable_popups'),
                               ('clean_pdf_links','clean_pdf_links')]:
                w = getattr(self, attr, None)
                if w and key in c: w.setChecked(c[key])
            self.update_style_preview()
        except Exception as e:
            print(f'Styles load error: {e}')

    def save_styles_config(self):
        cp = os.path.join(get_app_path(), 'styles_config.json')
        try:
            config = {}
            for attr, key in [
                ('pattern_combo','pattern'), ('color_combo','rect_color'),
                ('line_style_combo','line_style'), ('blink_speed_combo','blink_speed'),
                ('fill_combo','fill_style'), ('animation_combo','animation_type'),
                ('fill_color_combo','fill_color'), ('effect_combo','effect')]:
                w = getattr(self, attr, None)
                if w: config[key] = w.currentText()
            for attr, key in [
                ('line_width_spinbox','line_width'), ('duration_spinbox','duration'),
                ('opacity_spinbox','opacity'), ('corner_radius_spinbox','corner_radius'),
                ('rect_margin_spinbox','rect_margin')]:
                w = getattr(self, attr, None)
                if w: config[key] = w.value()
            config['custom_pattern'] = self.custom_pattern_input.text()
            for attr, key in [('keep_original_name','keep_original_name'),
                               ('disable_popups','disable_popups'),
                               ('clean_pdf_links','clean_pdf_links')]:
                w = getattr(self, attr, None)
                if w: config[key] = w.isChecked()
            with open(cp, 'w') as f: json.dump(config, f, indent=2)
        except Exception as e:
            print(f'Styles save error: {e}')

    # ─── PATTERN ──────────────────────────────
    def on_pattern_changed(self, name):
        self.current_pattern = name
        if name == 'Custom':
            self.custom_pattern_input.setEnabled(True)
            self.pattern_example.setText('Write your format')
            self.pattern_hint.show(); self.pattern_preview_label.show()
            self.on_custom_pattern_changed(self.custom_pattern_input.text())
        else:
            self.custom_pattern_input.setEnabled(False)
            self.pattern_hint.hide(); self.pattern_preview_label.hide()
            self.pattern_example.setText(self.REFERENCE_PATTERNS.get(name, {}).get('example', ''))

    def on_custom_pattern_changed(self, text):
        self.custom_pattern = text
        if text:
            regex, example, valid = self.convert_simple_pattern_to_regex(text)
            if valid:
                self.pattern_preview_label.setText(f'✓  Detects: {example}')
                self.pattern_preview_label.setStyleSheet(f"color: {COLORS['success']}; font-size: 12px;")
            else:
                self.pattern_preview_label.setText(f'⚠  Pattern: {regex}')
                self.pattern_preview_label.setStyleSheet(f"color: {COLORS['warning']}; font-size: 12px;")
        else:
            self.pattern_preview_label.setText('')

    def convert_simple_pattern_to_regex(self, simple_pattern):
        has_ph = any(p in simple_pattern.upper() for p in ['{P}','{C}','{F}','{PAG}','{COL}','{FILA}'])
        if has_ph:
            regex = simple_pattern
            for ch in ['\\', '.', '^', '$', '*', '+', '?', '[', ']', '(', ')', '|']:
                regex = regex.replace(ch, '\\' + ch)
            regex = re.sub(r'\{P\}|\{PAG\}|\{PAGINA\}', r'(\\d+)', regex, flags=re.IGNORECASE)
            regex = re.sub(r'\{C\}|\{COL\}|\{COLUMNA\}', r'(\\d+)', regex, flags=re.IGNORECASE)
            regex = re.sub(r'\{F\}|\{FILA\}', r'([A-Z])', regex, flags=re.IGNORECASE)
            example = simple_pattern
            example = re.sub(r'\{P\}|\{PAG\}|\{PAGINA\}', '5', example, flags=re.IGNORECASE)
            example = re.sub(r'\{C\}|\{COL\}|\{COLUMNA\}', '3', example, flags=re.IGNORECASE)
            example = re.sub(r'\{F\}|\{FILA\}', 'A', example, flags=re.IGNORECASE)
            return regex, example, True
        return simple_pattern, simple_pattern, False

    def get_current_pattern(self):
        if self.current_pattern == 'Custom':
            if not self.custom_pattern: return None
            regex, _, _ = self.convert_simple_pattern_to_regex(self.custom_pattern)
            return regex
        return self.REFERENCE_PATTERNS.get(self.current_pattern, {}).get('pattern', '')

    def get_pattern_groups_order(self):
        if self.current_pattern == 'Custom' and self.custom_pattern:
            pu = self.custom_pattern.upper()
            pm = re.search(r'\{P\}|\{PAG\}|\{PAGINA\}', pu)
            cm = re.search(r'\{C\}|\{COL\}|\{COLUMNA\}', pu)
            fm = re.search(r'\{F\}|\{FILA\}', pu)
            phs = sorted([(m.start(), n) for m, n in [(pm,'página'),(cm,'columna'),(fm,'fila')] if m], key=lambda x: x[0])
            groups = tuple(p[1] for p in phs)
            return groups if groups else ('página','columna','fila')
        return self.REFERENCE_PATTERNS.get(self.current_pattern, {}).get('groups', ('página','columna','fila'))

    # ─── STYLE HELPERS ────────────────────────
    def update_style_preview(self, *args):
        # Obtener valores actuales
        color_name = self.color_combo.currentText()
        line_width = self.line_width_spinbox.value()
        line_style = self.line_style_combo.currentText()
        opacity = self.opacity_spinbox.value()
        fill_style = self.fill_combo.currentText()
        fill_color_name = self.fill_color_combo.currentText()
        corner_radius = self.corner_radius_spinbox.value()
        
        # Mapeo de colores
        cmap = {'Red':'#EF4444','Green':'#22C55E','Blue':'#3B82F6','Yellow':'#FCD34D',
                'Orange':'#F97316','Magenta':'#D946EF','Cyan':'#06B6D4','White':'#FFFFFF','Black':'#000000'}
        border_color = cmap.get(color_name, '#EF4444')
        
        # Texto del preview
        preview_text = "REF-001"
        
        # Determinar estilo de borde
        border_style = "solid"
        if line_style == 'Dashed':
            border_style = "dashed"
        elif line_style == 'Dotted':
            border_style = "dotted"
        
        # Calcular tamaño de fuente basado en el ancho de línea
        font_size = 16 + (line_width * 2)  # Base 16px + 2px por cada unidad de ancho
        
        # Aplicar opacidad al color del borde
        opacity_value = int(opacity * 2.55)
        border_rgba = f"rgba({int(border_color[1:3], 16)}, {int(border_color[3:5], 16)}, {int(border_color[5:7], 16)}, {opacity_value/255:.2f})"
        
        # Determinar color de fondo
        background_style = "transparent"
        text_color = border_rgba
        
        if fill_style != 'None':
            # Determinar color de relleno
            if fill_color_name == 'Same as border':
                fill_color = border_color
            else:
                fill_color = cmap.get(fill_color_name, border_color)
            
            if fill_style == 'Semi-transparent':
                bg_opacity = opacity_value // 3  # 33% de la opacidad del borde
                background_style = f"rgba({int(fill_color[1:3], 16)}, {int(fill_color[3:5], 16)}, {int(fill_color[5:7], 16)}, {bg_opacity/255:.2f})"
            elif fill_style == 'Solid':
                background_style = f"rgba({int(fill_color[1:3], 16)}, {int(fill_color[3:5], 16)}, {int(fill_color[5:7], 16)}, {opacity_value/255:.2f})"
                # Si el fondo es sólido, ajustar color del texto para contraste
                if fill_color in ['#000000', '#EF4444', '#FF6B2B']:
                    text_color = "#FFFFFF"
                elif fill_color == '#FFFFFF':
                    text_color = "#000000"
        
        # Aplicar radio de esquinas
        border_radius = f"{corner_radius}px" if corner_radius > 0 else "4px"
        
        self.style_preview.setText(preview_text)
        self.style_preview.setStyleSheet(f"""
            QLabel {{
                color: {text_color}; 
                font-size: {font_size}px; 
                font-weight: 700;
                letter-spacing: 1px; 
                background: {background_style};
                border-radius: {border_radius};
                padding: 12px 24px;
                border: {line_width}px {border_style} {border_rgba};
                min-width: 100px;
                max-width: 200px;
            }}
        """)
        self.save_styles_config()

    def get_highlight_color(self):
        return {
            'Red':'color.red','Green':'color.green','Blue':'color.blue',
            'Yellow':'color.yellow','Orange':'["RGB", 0.976, 0.451, 0.086]',
            'Magenta':'color.magenta','Cyan':'color.cyan'
        }.get(self.color_combo.currentText(), 'color.red')

    def get_fill_color(self):
        n = self.fill_color_combo.currentText()
        if n == 'Same as border': return self.get_highlight_color()
        return {
            'Red':'color.red','Green':'color.green','Blue':'color.blue','Yellow':'color.yellow',
            'Orange':'["RGB", 0.976, 0.451, 0.086]','Magenta':'color.magenta','Cyan':'color.cyan',
            'White':'color.white','Black':'color.black'
        }.get(n, 'color.red')

    def get_blink_speed(self):
        return {'Fast':300,'Normal':500,'Slow':800,'None':0}.get(self.blink_speed_combo.currentText(), 500)

    def get_highlight_duration(self):
        return self.duration_spinbox.value() * 1000

    def update_rows_info(self, value):
        last = chr(ord('A') + value - 1) if value <= 26 else 'Z+'
        self.rows_info_label.setText(f'A–{last}')

    def update_size_placeholders(self):
        c, r = self.cols_spinbox.value(), self.rows_spinbox.value()
        self.col_sizes_input.setPlaceholderText(f'e.g. {",".join(["1"]*c)}')
        self.row_sizes_input.setPlaceholderText(f'e.g. {",".join(["1"]*r)}')

    # ─── DIALOGS ──────────────────────────────
    def show_pattern_help(self):
        msg = QMessageBox(self)
        msg.setWindowTitle('Pattern Help')
        msg.setTextFormat(Qt.RichText)
        msg.setText("""
<h3 style="color:#FF6B2B">Placeholders</h3>
<table cellpadding="6" style="border-collapse:collapse">
<tr><td><code>{P}</code></td><td>Page number</td><td>1, 25, 100…</td></tr>
<tr><td><code>{C}</code></td><td>Column number</td><td>0, 5, 12…</td></tr>
<tr><td><code>{F}</code></td><td>Row letter</td><td>A, B, Z…</td></tr>
</table>
<h3 style="color:#FF6B2B">Examples</h3>
<table cellpadding="6">
<tr><td><code>/{P}.{C}-{F}</code></td><td>→ /5.3-A</td></tr>
<tr><td><code>{P}-{F}.{C}</code></td><td>→ 25-A.0</td></tr>
<tr><td><code>[{P}/{F}/{C}]</code></td><td>→ [5/A/3]</td></tr>
</table>
""")
        msg.exec_()

    def show_references_dialog(self):
        d = QDialog(self)
        d.setWindowTitle('References')
        d.setMinimumSize(920, 600)
        d.setStyleSheet(f"QDialog {{ background: {COLORS['bg_base']}; }}")
        lay = QVBoxLayout(d)
        lay.setContentsMargins(20, 20, 20, 20); lay.setSpacing(16)

        hdr = QHBoxLayout()
        t = QLabel('References')
        t.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 18px; font-weight: 700;")
        hdr.addWidget(t); hdr.addStretch()
        hdr.addWidget(Badge(f'{len(self.references)} found', 'accent' if self.references else 'muted'))
        lay.addLayout(hdr)

        tbl = QTableWidget()
        multi = len(self.pdf_paths) > 1
        cols = (['PDF'] if multi else []) + ['Reference','Page','Column','Row','Context']
        tbl.setColumnCount(len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.setRowCount(len(self.references))
        tbl.setAlternatingRowColors(True)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(len(cols)-1, QHeaderView.Stretch)

        for row, ref in enumerate(self.references):
            off = 0
            if multi:
                n = ref.get('pdf_name','?')
                tbl.setItem(row, 0, QTableWidgetItem(n[:25]+'…' if len(n)>25 else n))
                off = 1
            rt = ref['full'] + (f" (#{ref['instance']})" if ref.get('instance',1)>1 else '')
            tbl.setItem(row, off, QTableWidgetItem(rt))
            tbl.setItem(row, off+1, QTableWidgetItem(ref['page']))
            tbl.setItem(row, off+2, QTableWidgetItem(ref['column']))
            tbl.setItem(row, off+3, QTableWidgetItem(ref['row']))
            ctx = ref['context'] + (f" [p.{ref['pdf_page']+1}]" if ref.get('coordinates') else '')
            tbl.setItem(row, off+4, QTableWidgetItem(ctx))
        lay.addWidget(tbl)

        close = make_btn('Close', 'accent')
        close.clicked.connect(d.accept)
        lay.addWidget(close, 0, Qt.AlignRight)
        d.exec_()

    def show_statistics_dialog(self):
        d = QDialog(self)
        d.setWindowTitle('Statistics')
        d.setMinimumSize(560, 460)
        d.setStyleSheet(f"QDialog {{ background: {COLORS['bg_base']}; }}")
        lay = QVBoxLayout(d)
        lay.setContentsMargins(20, 20, 20, 20); lay.setSpacing(16)
        t = QLabel('Analysis Statistics')
        t.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 18px; font-weight: 700;")
        lay.addWidget(t)
        sd = QTextEdit()
        sd.setReadOnly(True)
        sd.setText(self.stats_text.toPlainText())
        lay.addWidget(sd)
        close = make_btn('Close', 'accent')
        close.clicked.connect(d.accept)
        lay.addWidget(close, 0, Qt.AlignRight)
        d.exec_()

    # ─── GRID EDITOR ──────────────────────────
    def open_visual_editor(self):
        if not self.pdf_path:
            self.show_warning_dialog('No PDF Loaded', 'Load a PDF first.')
            return
        dlg = GridEditorDialog(self, self.pdf_path)
        if dlg.exec_() == QDialog.Accepted:
            gd = dlg.get_grid_data()
            cp, rp = gd['column_positions'], gd['row_positions']
            if len(cp) < 2: 
                self.show_warning_dialog('Insufficient Grid Lines', 'Need at least 2 column lines.')
                return
            if len(rp) < 2: 
                self.show_warning_dialog('Insufficient Grid Lines', 'Need at least 2 row lines.')
                return
            self.column_positions, self.row_positions = cp, rp
            self.grid_detected = True
            nc, nr = len(cp)-1, len(rp)-1
            self.cols_spinbox.setValue(nc); self.rows_spinbox.setValue(nr)
            pw, ph = gd['page_width'], gd['page_height']
            
            # Guardar dimensiones de página para plantillas
            self.grid_page_width = pw
            self.grid_page_height = ph
            
            if cp: self.margin_left_spinbox.setValue(max(0, min(int(cp[0]/pw*100), 30)))
            if rp: self.margin_top_spinbox.setValue(max(0, min(int(rp[0]/ph*100), 30)))
            self._compute_relative_sizes()
            
            # Guardar configuración
            self.save_grid_config()
            
            self.grid_template_btn.setText(f'● Grid {nc}×{nr}')
            self.grid_template_btn.setStyleSheet(f"""
                QPushButton {{
                    color: {COLORS['accent']};
                    font-size: 11px;
                    font-weight: 700;
                    padding: 8px 6px;
                    background: {COLORS['accent_dim']};
                    border: 1px solid {COLORS['accent']};
                    border-radius: 8px;
                    text-align: center;
                }}
                QPushButton:hover {{
                    background: {COLORS['accent_hover']};
                    color: {COLORS['bg_base']};
                }}
            """)

    # ─── AUTODETECT ───────────────────────────
    def autodetect_grid(self):
        pass  # Keep for compat

    def filter_close_lines(self, lines, min_distance):
        if not lines: return []
        filtered = [lines[0]]
        for line in lines[1:]:
            if line - filtered[-1] >= min_distance: filtered.append(line)
        return filtered

    def calculate_relative_sizes(self, positions):
        if len(positions) < 2: return []
        dists = [positions[i+1]-positions[i] for i in range(len(positions)-1)]
        if not dists: return []
        mn = min(dists)
        if mn <= 0: return [1.0]*len(dists)
        return [d/mn for d in dists]

    def update_size_placeholders(self):
        c, r = self.cols_spinbox.value(), self.rows_spinbox.value()
        self.col_sizes_input.setPlaceholderText(f'e.g. {",".join(["1"]*c)}')
        self.row_sizes_input.setPlaceholderText(f'e.g. {",".join(["1"]*r)}')

    def parse_sizes(self, text, count):
        if not text.strip(): return [1.0]*count
        try:
            s = [float(x.strip()) for x in text.split(',') if x.strip()]
            while len(s) < count: s.append(1.0)
            return s[:count]
        except ValueError:
            return [1.0]*count

    # ─── DETECTION ────────────────────────────
    def detect_references(self):
        """Detección de referencias con threading y barra integrada"""
        if not self.pdf_paths: 
            self.show_warning_dialog('No PDFs Loaded', 'No PDFs loaded.')
            return
        pattern = self.get_current_pattern()
        if not pattern: 
            self.show_warning_dialog('No Pattern', 'Enter a valid regex pattern.')
            return
        try: 
            re.compile(pattern)
        except re.error as e: 
            self.show_error_dialog('Invalid Regex', f'Invalid regex pattern:\n{e}')
            return

        # Mostrar barra de progreso en status bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage('Starting detection...')
        
        # Deshabilitar botones durante el procesamiento
        self.detect_button.setEnabled(False)
        self.generate_button.setEnabled(False)
        self.stats_button.setEnabled(False)
        
        # Obtener orden de grupos
        groups_order = self.get_pattern_groups_order()
        
        # Crear y configurar worker thread
        self.detection_worker = DetectionWorker(self.pdf_paths, pattern, groups_order)
        self.detection_worker.progress_updated.connect(self.on_detection_progress)
        self.detection_worker.finished_signal.connect(self.on_detection_finished)
        self.detection_worker.error_signal.connect(self.on_detection_error)
        
        # Iniciar el worker thread
        self.detection_worker.start()

    def populate_table(self):
        multi = len(self.pdf_paths) > 1
        cols = (['PDF'] if multi else []) + ['Reference','Page','Column','Row','Context']
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setRowCount(len(self.references))
        for row, ref in enumerate(self.references):
            off = 0
            if multi: self.table.setItem(row,0,QTableWidgetItem(ref.get('pdf_name','?')[:25])); off=1
            rt = ref['full'] + (f" (#{ref['instance']})" if ref.get('instance',1)>1 else '')
            self.table.setItem(row,off,QTableWidgetItem(rt))
            self.table.setItem(row,off+1,QTableWidgetItem(ref['page']))
            self.table.setItem(row,off+2,QTableWidgetItem(ref['column']))
            self.table.setItem(row,off+3,QTableWidgetItem(ref['row']))
            ctx = ref['context']+(f" [p.{ref['pdf_page']+1}]" if ref.get('coordinates') else '')
            self.table.setItem(row,off+4,QTableWidgetItem(ctx))

    def update_statistics(self, total):
        pi = self.REFERENCE_PATTERNS.get(self.current_pattern, {})
        if total == 0:
            self.stats_text.setText(f"Pattern: {self.current_pattern}\nNo references found.")
        else:
            unique = len(set(r['full'] for r in self.references))
            pages = len(set(r['page'] for r in self.references))
            pc = {}
            for r in self.references: pc[r['page']] = pc.get(r['page'],0)+1
            lines = [f"Pattern: {self.current_pattern}","",
                     f"Total: {total}","Unique: {unique}",f"Pages referenced: {pages}","",
                     "By page:"]
            for p in sorted(pc, key=lambda x: int(x) if x.isdigit() else 0):
                lines.append(f"  Page {p}: {pc[p]}")
            self.stats_text.setText('\n'.join(lines))
        self.ref_count_label.setText(str(total))

    # ─── PDF GENERATION ───────────────────────
    def get_javascript_code(self):
        color = self.get_highlight_color()
        lw = self.line_width_spinbox.value()
        bs = self.get_blink_speed()
        dur = self.get_highlight_duration()
        fill_style = self.fill_combo.currentText()
        fill_color = self.get_fill_color()
        margin = self.rect_margin_spinbox.value()
        animation = self.animation_combo.currentText()
        opacity = self.opacity_spinbox.value() / 100.0
        line_style = self.line_style_combo.currentText()

        # Configurar relleno
        fill_code = ''
        if fill_style in ('Semi-transparent', 'Solid'): 
            fill_code = f"f.fillColor = {fill_color};"

        # Configurar márgenes
        margin_code = ''
        if margin != 0:
            margin_code = f"""
    coordinates[0] -= {margin};
    coordinates[1] -= {margin};
    coordinates[2] += {margin};
    coordinates[3] += {margin};"""

        # Configurar estilo de línea
        line_style_code = ""
        if line_style == 'Dashed':
            line_style_code = "f.borderStyle = border.d;"
        elif line_style == 'Dotted':
            line_style_code = "f.borderStyle = border.d;"

        # Configurar animación - usar la lógica exacta del antiguo
        if animation == 'None' or bs == 0:
            blinker_code = "// Sin animación"
            interval_code = ""
        else:
            blinker_code = """var f = getField('Target');
    if (f != null) {
        var oldDirty = dirty;
        if (interval.counter++%2) { f.hidden=false; }
        else { f.hidden = true; }
        dirty = oldDirty;
    }"""
            interval_code = f"interval = app.setInterval('blinker()', {bs});\n    interval.counter = 0;"

        # Usar exactamente la misma estructura que el antiguo.py
        return f"""
function finish() {{
    app.clearInterval(interval);
    var oldDirty = dirty;
    removeField('Target');
    dirty = oldDirty;
}}

function blinker() {{
    {blinker_code}
}}

function highlight(page, coordinates) {{
    var f = getField('Target');
    if (f != null) {{
        app.clearTimeOut(timer);
        finish();
    }}{margin_code}
    var oldDirty = dirty;
    var f = addField('Target', 'button', page, coordinates);
    f.lineWidth = {lw};
    f.strokeColor = {color};
    {fill_code}
    {line_style_code}
    dirty = oldDirty;
    {interval_code}
    timer = app.setTimeOut('finish()', {dur});
}}
"""


    def _copy_bookmarks(self, reader, writer):
        try:
            if not hasattr(reader,'outline') or not reader.outline: return
            page_map = {}
            for i, page in enumerate(reader.pages):
                if hasattr(page,'indirect_reference'):
                    ref = page.indirect_reference
                    if hasattr(ref,'idnum'): page_map[ref.idnum] = i
                    else: page_map[id(ref)] = i
                else: page_map[id(page)] = i

            def proc(items, parent=None):
                last = parent
                for item in items:
                    if isinstance(item, list): proc(item, last)
                    else:
                        try:
                            title = item.get('/Title','Untitled')
                            if hasattr(title,'get_object'): title = title.get_object()
                            pn = None
                            for key in ('/Page', '/Dest', '/A'):
                                if pn is not None: break
                                if key in item:
                                    obj = item[key]
                                    if hasattr(obj,'get_object'): obj = obj.get_object()
                                    if key == '/A' and isinstance(obj,dict) and '/D' in obj:
                                        obj = obj['/D']
                                        if hasattr(obj,'get_object'): obj = obj.get_object()
                                    target = obj if key == '/Page' else (obj[0] if isinstance(obj,list) and obj else None)
                                    if target is None: continue
                                    if hasattr(target,'idnum') and target.idnum in page_map: pn = page_map[target.idnum]
                                    elif hasattr(target,'indirect_reference'):
                                        r = target.indirect_reference
                                        if hasattr(r,'idnum') and r.idnum in page_map: pn = page_map[r.idnum]
                            if pn is not None:
                                last = writer.add_outline_item(str(title), pn, parent)
                        except Exception as e: print(f'Bookmark warning: {e}')
                return last
            proc(reader.outline)
        except Exception as e: print(f'Bookmarks copy error: {e}')

    def optimize_pdf_file(self, path):
        try:
            import tempfile, shutil
            fd, tmp = tempfile.mkstemp(suffix='.pdf'); os.close(fd)
            doc = fitz.open(path)
            doc.save(tmp, deflate=True, clean=True, garbage=3, linear=True, pretty=False)
            doc.close(); shutil.move(tmp, path)
            return path
        except Exception as e: print(f'Optimize error: {e}'); return path

    def generate_interactive_pdf(self):
        if not self.all_references: 
            self.show_warning_dialog('No References', 'Detect references first.')
            return
        keep = self.keep_original_name.isChecked()
        if keep:
            if QMessageBox.question(self,'Overwrite?',
                f'This will overwrite {len(self.pdf_paths)} original file(s). Continue?',
                QMessageBox.Yes|QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
            output_dir = None
        else:
            if len(self.pdf_paths) == 1:
                sp, _ = QFileDialog.getSaveFileName(self,'Save Interactive PDF',
                    self.pdf_paths[0].replace('.pdf','_interactive.pdf'),'PDF Files (*.pdf)')
                if not sp: return
                output_dir = os.path.dirname(sp)
            else:
                output_dir = QFileDialog.getExistingDirectory(self,'Select output folder',
                    os.path.dirname(self.pdf_paths[0]))
                if not output_dir: return

        # Mostrar barra de progreso en status bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage('Starting PDF generation...')
        
        # Deshabilitar botones durante el procesamiento
        self.detect_button.setEnabled(False)
        self.generate_button.setEnabled(False)
        self.stats_button.setEnabled(False)
        
        # Crear y configurar worker thread
        self.pdf_worker = PDFGenerationWorker(self, self.all_references, keep, output_dir)
        self.pdf_worker.progress_updated.connect(self.on_pdf_progress)
        self.pdf_worker.finished_signal.connect(self.on_pdf_finished)
        self.pdf_worker.error_signal.connect(self.on_pdf_error)
        
        # Iniciar el worker thread
        self.pdf_worker.start()

    def on_pdf_progress(self, value, text):
        """Actualizar progreso de generación de PDF"""
        self.progress_bar.setValue(value)
        self.statusBar().showMessage(text)
        self.statusBar().showMessage(text)
        
    def on_pdf_finished(self, generated_files, total_links):
        """Manejar finalización de generación de PDF"""
        # Ocultar barra de progreso
        self.progress_bar.setVisible(False)
        
        # Rehabilitar botones
        self.detect_button.setEnabled(True)
        self.generate_button.setEnabled(True)
        self.stats_button.setEnabled(True)
        
        # Mostrar diálogo de éxito si no está deshabilitado
        if not self.disable_popups.isChecked() and generated_files:
            n = len(generated_files)
            msg_title = f'{"PDF" if n==1 else str(n)+" PDFs"} generated successfully!'
            msg_links = f'Links added: {total_links}'
            if n == 1:
                msg_file = f'File: {generated_files[0]}'
            else:
                msg_file = f'Folder: {os.path.dirname(generated_files[0])}'
            
            self.show_success_dialog(msg_title, msg_links, msg_file)
        
        self.statusBar().showMessage(f'{len(generated_files)} PDF(s) generated  ·  {total_links} links')
        
    def on_pdf_error(self, error_msg):
        """Manejar errores de generación de PDF"""
        self.progress_bar.setVisible(False)
        self.detect_button.setEnabled(True)
        self.generate_button.setEnabled(True)
        self.stats_button.setEnabled(True)
        self.statusBar().showMessage('Error during PDF generation')
        self.show_error_dialog('PDF Generation Error', error_msg)

    def _generate_single_pdf(self, pdf_path, references, keep_original_name, output_dir):
        """Generar un PDF individual con referencias interactivas"""
        if not references:
            return None
            
        try:
            # Determinar archivo de salida
            if keep_original_name:
                output_path = pdf_path + '.tmp'
                final_path = pdf_path
            else:
                output_path = os.path.join(output_dir, os.path.basename(pdf_path).replace('.pdf', '_interactive.pdf'))
                final_path = output_path
            
            # Optimizar PDF si está habilitado
            working_pdf = pdf_path
            if self.optimize_pdf.isChecked():
                working_pdf = self.optimize_pdf_file(pdf_path)
            
            # Procesar referencias con PyMuPDF para obtener coordenadas
            tmp_doc = fitz.open(working_pdf)
            refs_data = []
            
            for ref in references:
                if not ref['coordinates']:
                    continue
                    
                sp = tmp_doc[ref['pdf_page']]
                x0, y0, x1, y1 = ref['coordinates']
                ph = sp.rect.height
                tpn = int(ref['page']) - 1
                
                if tpn < 0 or tpn >= len(tmp_doc):
                    continue
                    
                tc = self.calculate_target_coordinates(tmp_doc[tpn], ref['column'], ref['row'])
                tph = tmp_doc[tpn].rect.height
                
                refs_data.append({
                    'full': ref['full'],
                    'pdf_page': ref['pdf_page'],
                    'coordinates': [x0, ph-y1, x1, ph-y0],
                    'target_page': tpn,
                    'target_coordinates': [tc[0], tph-tc[3], tc[2], tph-tc[1]]
                })
            
            tmp_doc.close()
            
            # Crear PDF interactivo con PyPDF2
            reader = PdfReader(working_pdf)
            writer = PdfWriter()
            writer.append_pages_from_reader(reader)
            
            # Limpiar enlaces existentes si está habilitado
            if self.clean_pdf_links.isChecked():
                for page in writer.pages:
                    if "/Annots" in page:
                        cleaned = []
                        for ar in page["/Annots"]:
                            try:
                                a = ar.get_object() if hasattr(ar, 'get_object') else ar
                                if str(a.get("/Subtype", "")) != "/Link":
                                    cleaned.append(ar)
                            except:
                                pass
                        if cleaned:
                            page[NameObject("/Annots")] = ArrayObject(cleaned)
                        else:
                            del page["/Annots"]
            
            # Inyectar JavaScript
            writer.add_js(self.get_javascript_code())
            
            # Crear anotaciones de enlaces
            for rd in refs_data:
                try:
                    page = writer.pages[rd['pdf_page']]
                    coords = rd['coordinates']
                    tpage = rd['target_page']
                    tc = rd['target_coordinates']
                    
                    # Crear acción JavaScript
                    js_call = f"highlight({tpage},{tc});"
                    print(f"=== LLAMADA JS: {js_call} ===")
                    js_act = DictionaryObject({
                        NameObject("/S"): NameObject("/JavaScript"),
                        NameObject("/JS"): create_string_object(js_call)
                    })
                    
                    # Crear acción GoTo
                    goto_act = DictionaryObject({
                        NameObject("/S"): NameObject("/GoTo"),
                        NameObject("/D"): ArrayObject([
                            writer.pages[tpage].indirect_reference,
                            NameObject("/XYZ"),
                            NumberObject(int(tc[0])),
                            NumberObject(int(tc[3])),
                            NumberObject(0)
                        ])
                    })
                    goto_act[NameObject("/Next")] = js_act
                    
                    # Crear anotación de enlace
                    annot = DictionaryObject({
                        NameObject("/Type"): NameObject("/Annot"),
                        NameObject("/Subtype"): NameObject("/Link"),
                        NameObject("/Rect"): ArrayObject([NumberObject(c) for c in coords]),
                        NameObject("/Border"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(0)]),
                        NameObject("/A"): goto_act,
                        NameObject("/H"): NameObject("/N")
                    })
                    
                    if "/Annots" in page:
                        page["/Annots"].append(annot)
                    else:
                        page[NameObject("/Annots")] = ArrayObject([annot])
                        
                except Exception as re_:
                    print(f'Reference error: {re_}')
            
            # Copiar marcadores si existen
            try:
                if hasattr(reader, 'outline') and reader.outline:
                    self._copy_bookmarks(reader, writer)
            except Exception as be:
                print(f'Bookmark error: {be}')
            
            # Escribir archivo
            with open(output_path, 'wb') as fh:
                writer.write(fh)
            
            # Mover archivo si es necesario
            if keep_original_name and output_path != final_path:
                import shutil
                shutil.move(output_path, final_path)
            
            return final_path
            
        except Exception as e:
            print(f'PDF generation error for {pdf_path}: {e}')
            return None

    def calculate_target_coordinates(self, target_page, column, row):
        rect = target_page.rect; width, height = rect.width, rect.height
        try: col_num = int(column)
        except (ValueError, TypeError):
            col_num = ord(column.upper())-ord('A') if column and str(column).isalpha() else 0
        row_index = 0
        if row:
            if str(row).isalpha():
                row_index = ord(row[0].upper())-ord('A') if len(row)==1 else 0
            elif str(row).isdigit():
                row_index = int(row)

        if self.grid_detected and self.column_positions and self.row_positions:
            ci = max(0, min(col_num, len(self.column_positions)-2))
            ri = max(0, min(row_index, len(self.row_positions)-2))
            x0 = self.column_positions[ci]
            x1 = self.column_positions[ci+1] if ci+1 < len(self.column_positions) else x0+50
            y0 = self.row_positions[ri]
            y1 = self.row_positions[ri+1] if ri+1 < len(self.row_positions) else y0+50
            return [max(0,min(x0,width)),max(0,min(y0,height)),max(0,min(x1,width)),max(0,min(y1,height))]

        ml = width*self.margin_left_spinbox.value()/100
        mt = height*self.margin_top_spinbox.value()/100
        uw, uh = width-2*ml, height-2*mt
        cp, rp = self.cols_spinbox.value(), self.rows_spinbox.value()
        cs = self.parse_sizes(self.col_sizes_input.text(), cp)
        rs = self.parse_sizes(self.row_sizes_input.text(), rp)
        col_num = max(0, min(col_num, cp-1)); row_index = max(0, min(row_index, rp-1))
        cu, ru = uw/sum(cs), uh/sum(rs)
        x0 = ml + sum(cs[i]*cu for i in range(col_num))
        x1 = x0 + cs[col_num]*cu
        y0 = mt + sum(rs[i]*ru for i in range(row_index))
        y1 = y0 + rs[row_index]*ru
        return [max(0,min(x0,width)),max(0,min(y0,height)),max(0,min(x1,width)),max(0,min(y1,height))]

    def on_detection_progress(self, value, text):
        """Actualizar progreso de detección"""
        self.progress_bar.setValue(value)
        self.statusBar().showMessage(text)
        
    def on_detection_finished(self, references, all_references):
        """Manejar finalización de detección"""
        self.references = references
        self.all_references = all_references
        
        # Ocultar barra de progreso
        self.progress_bar.setVisible(False)
        
        # Actualizar UI
        self.populate_table()
        self.update_statistics(len(self.references))
        
        if self.references: 
            self.generate_button.setEnabled(True)
            self.stats_button.setEnabled(True)

        self.ref_badge.setText(f'{len(self.references)} refs')
        self.statusBar().showMessage(f'Done  ·  {len(self.references)} references found in {len(self.pdf_paths)} file(s)')
        
        # Rehabilitar botón detect
        self.detect_button.setEnabled(True)
        
    def on_detection_error(self, error_msg):
        """Manejar errores de detección"""
        self.progress_bar.setVisible(False)
        self.detect_button.setEnabled(True)
        self.statusBar().showMessage('Error during detection')
        self.show_error_dialog('Detection Error', f'Detection error:\n{error_msg}')

    def show_success_dialog(self, title, links_info, file_info):
        """Mostrar diálogo de éxito personalizado completamente sin bordes"""
        dialog = QDialog(self)
        dialog.setWindowTitle('Done')
        dialog.setFixedSize(520, 380)
        dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dialog.setAttribute(Qt.WA_TranslucentBackground)
        
        # Layout principal
        main_layout = QVBoxLayout(dialog)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Contenedor principal sin bordes
        container = QWidget()
        container.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['bg_base']};
                border: 2px solid {COLORS['accent']};
                border-radius: 28px;
            }}
        """)
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(45, 35, 45, 35)
        layout.setSpacing(25)
        
        # Barra de título minimalista
        title_bar = QWidget()
        title_bar.setFixedHeight(40)
        title_bar.setStyleSheet("background: transparent; border: none;")
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        # Icono de éxito
        icon_label = QLabel()
        icon_label.setFixedSize(36, 36)
        icon_label.setStyleSheet(f"""
            QLabel {{
                background-color: {COLORS['accent']};
                border: none;
                border-radius: 18px;
                color: white;
                font-size: 20px;
                font-weight: bold;
            }}
        """)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setText('✓')
        
        # Botón cerrar minimalista
        close_btn = QPushButton('×')
        close_btn.setFixedSize(28, 28)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 18px;
                font-weight: bold;
                border-radius: 14px;
            }}
            QPushButton:hover {{
                background-color: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
            }}
        """)
        close_btn.clicked.connect(dialog.accept)
        
        title_layout.addWidget(icon_label)
        title_layout.addStretch()
        title_layout.addWidget(close_btn)
        
        layout.addWidget(title_bar)
        
        # Contenido principal - completamente sin bordes
        content_layout = QVBoxLayout()
        content_layout.setSpacing(20)
        content_layout.setContentsMargins(0, 5, 0, 0)
        
        # Título principal
        main_title = QLabel(title)
        main_title.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 20px;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
        """)
        main_title.setWordWrap(True)
        main_title.setAlignment(Qt.AlignCenter)
        
        # Información de enlaces
        links_label = QLabel(links_info)
        links_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['accent']};
                font-size: 16px;
                font-weight: 600;
                background: transparent;
                border: none;
            }}
        """)
        links_label.setAlignment(Qt.AlignCenter)
        
        # Información de archivo
        file_label = QLabel(file_info)
        file_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_secondary']};
                font-size: 13px;
                background: transparent;
                border: none;
                padding: 0px;
            }}
        """)
        file_label.setWordWrap(True)
        file_label.setAlignment(Qt.AlignCenter)
        
        # Sección de instrucciones - sin bordes visibles
        instructions_container = QWidget()
        instructions_container.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['bg_surface']};
                border: none;
                border-radius: 16px;
            }}
        """)
        
        instructions_layout = QVBoxLayout(instructions_container)
        instructions_layout.setSpacing(6)
        instructions_layout.setContentsMargins(20, 16, 20, 16)
        
        # Título de instrucciones
        inst_title = QLabel('To enable animations in Adobe Reader:')
        inst_title.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 13px;
                font-weight: 600;
                background: transparent;
                border: none;
                margin-bottom: 6px;
            }}
        """)
        inst_title.setAlignment(Qt.AlignCenter)
        
        # Lista de instrucciones
        instructions = [
            '• Open PDF in Adobe Acrobat Reader',
            '• Enable JavaScript in Preferences',
            '• Click links to see animations!'
        ]
        
        instructions_layout.addWidget(inst_title)
        
        for instruction in instructions:
            inst_label = QLabel(instruction)
            inst_label.setStyleSheet(f"""
                QLabel {{
                    color: {COLORS['text_secondary']};
                    font-size: 12px;
                    background: transparent;
                    border: none;
                    padding: 1px 0px;
                }}
            """)
            inst_label.setAlignment(Qt.AlignCenter)
            instructions_layout.addWidget(inst_label)
        
        # Botón OK
        ok_button = QPushButton('OK')
        ok_button.setFixedHeight(44)
        ok_button.setFixedWidth(100)
        ok_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['accent']};
                color: white;
                border: none;
                border-radius: 22px;
                font-size: 15px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {COLORS['accent_hover']};
            }}
            QPushButton:pressed {{
                background-color: {COLORS['accent_dim']};
            }}
        """)
        ok_button.clicked.connect(dialog.accept)
        
        # Agregar widgets al layout de contenido
        content_layout.addWidget(main_title)
        content_layout.addWidget(links_label)
        content_layout.addWidget(file_label)
        content_layout.addWidget(instructions_container)
        
        # Layout para el botón OK centrado
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addStretch()
        
        layout.addLayout(content_layout)
        layout.addLayout(button_layout)
        
        main_layout.addWidget(container)
        
        # Centrar el diálogo en la ventana principal
        dialog.adjustSize()
        parent_rect = self.frameGeometry()
        dialog_rect = dialog.frameGeometry()
        center_point = parent_rect.center()
        dialog_rect.moveCenter(center_point)
        dialog.move(dialog_rect.topLeft())
        
        # Mostrar el diálogo
        dialog.exec_()

    def show_error_dialog(self, title, message):
        """Mostrar diálogo de error personalizado con el estilo de la ventana principal"""
        dialog = QDialog(self)
        dialog.setWindowTitle('Error')
        dialog.setFixedSize(480, 320)
        dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dialog.setAttribute(Qt.WA_TranslucentBackground)
        
        # Layout principal
        main_layout = QVBoxLayout(dialog)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Contenedor principal con bordes redondeados
        container = QWidget()
        container.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['bg_base']};
                border: 2px solid {COLORS['danger']};
                border-radius: 28px;
            }}
        """)
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(30, 25, 30, 25)
        layout.setSpacing(20)
        
        # Barra de título personalizada
        title_bar = QWidget()
        title_bar.setFixedHeight(40)
        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        # Icono y título
        icon_label = QLabel()
        icon_label.setFixedSize(32, 32)
        icon_label.setStyleSheet(f"""
            QLabel {{
                background-color: {COLORS['danger']};
                border-radius: 16px;
                color: white;
                font-size: 18px;
                font-weight: bold;
            }}
        """)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setText('✕')
        
        title_label = QLabel('Error')
        title_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 18px;
                font-weight: 600;
                margin-left: 10px;
            }}
        """)
        
        # Botón cerrar
        close_btn = QPushButton('×')
        close_btn.setFixedSize(32, 32)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: none;
                color: {COLORS['text_secondary']};
                font-size: 20px;
                font-weight: bold;
                border-radius: 16px;
            }}
            QPushButton:hover {{
                background-color: {COLORS['bg_hover']};
                color: {COLORS['text_primary']};
            }}
            QPushButton:pressed {{
                background-color: {COLORS['danger']};
                color: white;
            }}
        """)
        close_btn.clicked.connect(dialog.accept)
        
        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        title_layout.addWidget(close_btn)
        
        layout.addWidget(title_bar)
        
        # Contenido del mensaje
        message_label = QLabel(message)
        message_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 14px;
                line-height: 1.4;
                background-color: {COLORS['bg_surface']};
                border: 1px solid {COLORS['danger']};
                border-radius: 16px;
                padding: 20px;
            }}
        """)
        message_label.setWordWrap(True)
        message_label.setAlignment(Qt.AlignTop)
        
        # Botón OK
        ok_button = QPushButton('OK')
        ok_button.setFixedHeight(44)
        ok_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['danger']};
                color: white;
                border: none;
                border-radius: 18px;
                font-size: 14px;
                font-weight: 600;
                padding: 0 30px;
            }}
            QPushButton:hover {{
                background-color: #DC2626;
            }}
            QPushButton:pressed {{
                background-color: #B91C1C;
            }}
        """)
        ok_button.clicked.connect(dialog.accept)
        
        # Layout para el botón OK centrado
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addStretch()
        
        layout.addWidget(message_label)
        layout.addStretch()
        layout.addLayout(button_layout)
        
        main_layout.addWidget(container)
        
        # Centrar el diálogo en la ventana principal
        parent_geometry = self.geometry()
        dialog_geometry = dialog.geometry()
        x = parent_geometry.x() + (parent_geometry.width() - dialog.width()) // 2
        y = parent_geometry.y() + (parent_geometry.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        # Mostrar el diálogo
        dialog.exec_()

    def show_warning_dialog(self, title, message):
        """Mostrar diálogo de advertencia simple y minimalista"""
        dialog = QDialog(self)
        dialog.setWindowTitle('Warning')
        dialog.setFixedSize(320, 180)
        dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dialog.setAttribute(Qt.WA_TranslucentBackground)
        
        # Layout principal
        main_layout = QVBoxLayout(dialog)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Contenedor principal
        container = QWidget()
        container.setStyleSheet(f"""
            QWidget {{
                background-color: {COLORS['bg_base']};
                border: 2px solid {COLORS['accent']};
                border-radius: 24px;
            }}
        """)
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(30, 25, 30, 25)
        layout.setSpacing(20)
        
        # Icono y mensaje en una sola línea
        content_layout = QVBoxLayout()
        content_layout.setSpacing(15)
        content_layout.setAlignment(Qt.AlignCenter)
        
        # Icono simple
        icon_label = QLabel('⚠')
        icon_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['accent']};
                font-size: 32px;
                background: transparent;
                border: none;
            }}
        """)
        icon_label.setAlignment(Qt.AlignCenter)
        
        # Mensaje
        message_label = QLabel(message)
        message_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                font-size: 15px;
                background: transparent;
                border: none;
            }}
        """)
        message_label.setWordWrap(True)
        message_label.setAlignment(Qt.AlignCenter)
        
        # Botón OK
        ok_button = QPushButton('OK')
        ok_button.setFixedHeight(40)
        ok_button.setFixedWidth(80)
        ok_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['accent']};
                color: white;
                border: none;
                border-radius: 20px;
                font-size: 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {COLORS['accent_hover']};
            }}
            QPushButton:pressed {{
                background-color: {COLORS['accent_dim']};
            }}
        """)
        ok_button.clicked.connect(dialog.accept)
        
        # Agregar elementos
        content_layout.addWidget(icon_label)
        content_layout.addWidget(message_label)
        
        # Layout para el botón OK centrado
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addStretch()
        
        layout.addLayout(content_layout)
        layout.addLayout(button_layout)
        
        main_layout.addWidget(container)
        
        # Centrar el diálogo en la ventana principal
        parent_geometry = self.geometry()
        dialog_geometry = dialog.geometry()
        x = parent_geometry.x() + (parent_geometry.width() - dialog.width()) // 2
        y = parent_geometry.y() + (parent_geometry.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        # Mostrar el diálogo
        dialog.exec_()

    def coords_match(self, c1, c2, tol=5):
        return c1 and c2 and all(abs(c1[i]-c2[i])<=tol for i in range(4))


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    icon_path = os.path.join(get_app_path(), 'logo.png')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = PDFReferenceDetector()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()