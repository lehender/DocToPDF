# doctopdf.py
# deps: PySide6
# Office -> PDF via LibreOffice headless (bundled or system)
import sys, os, shutil, subprocess, platform
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QListWidget, QFrame, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon

def roots_to_search():
    """Return all plausible base folders for bundled resources."""
    roots = []
    # 1) Onedir: folder containing the EXE
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).parent)
    # 2) PyInstaller's temp dir (onedir: _internal, onefile: temp unpack dir)
    if hasattr(sys, "_MEIPASS") and sys._MEIPASS:
        roots.append(Path(sys._MEIPASS))
    # 3) Dev run: folder containing this .py
    roots.append(Path(__file__).parent)
    # De-dup while preserving order
    seen, unique = set(), []
    for r in roots:
        if r not in seen:
            unique.append(r); seen.add(r)
    return unique

def find_icon_path():
    for base in roots_to_search():
        for rel in [("assets","app.ico"), ("assets","app.png")]:
            p = base.joinpath(*rel)
            if p.exists():
                return str(p)
    return None

def find_soffice() -> str | None:
    portable_rel_paths = [
        ["tools","LibreOfficePortable","App","libreoffice","program","soffice.exe"],
        ["tools","libreoffice","App","libreoffice","program","soffice.exe"],
        ["tools","libreoffice","program","soffice.exe"],
        ["tools","LibreOffice","program","soffice.exe"],
    ]
    # mac/linux variants (ignored on Windows)
    if platform.system() == "Darwin":
        portable_rel_paths += [
            ["tools","LibreOffice.app","Contents","MacOS","soffice"],
            ["tools","libreoffice","program","soffice"],
        ]
    else:
        portable_rel_paths += [["tools","libreoffice","program","soffice"]]

    for base in roots_to_search():
        for parts in portable_rel_paths:
            p = base.joinpath(*parts)
            if p.exists():
                return str(p)

    # PATH fallback
    for name in ("soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        found = shutil.which(name) or (Path(name).exists() and name)
        if found:
            return str(found)
    return None

def convert_with_libreoffice(src: Path, dst: Path, soffice: str):
    dst.parent.mkdir(parents=True, exist_ok=True)
    outdir = dst.parent
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)]
    subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    produced = outdir / (src.stem + ".pdf")
    if produced != dst:
        if dst.exists():
            dst.unlink()
        produced.replace(dst)

def is_supported(p: Path):
    return p.suffix.lower() in {
        ".docx", ".doc", ".odt",
        ".pptx", ".ppt", ".odp",
        ".xlsx", ".xls", ".ods"
    }

# ---------------------------
# Styling
# ---------------------------
CARD_CSS = """
QFrame#Card {
    border: 2px solid #3f3f46;
    border-radius: 12px;
    background: #1f1f23;
}
QLabel#Title {
    color: #e5e7eb;
    font-size: 16px;
    font-weight: 600;
}
QLabel#Hint, QLabel#Subtle {
    color: #a1a1aa;
    font-size: 12px;
}
QPushButton#Primary {
    border: none;
    border-radius: 10px;
    padding: 10px 14px;
    background: #2563eb;
    color: white;
    font-weight: 600;
}
QPushButton#Primary:hover { background: #1d4ed8; }
QPushButton#Primary:pressed { background: #1e40af; }

QPushButton#Ghost {
    border: 1px solid #3f3f46;
    border-radius: 10px;
    padding: 8px 12px;
    background: #18181b;
    color: #e5e7eb;
}
QPushButton#Ghost:hover { background: #111113; }

QListWidget {
    border: 2px solid #3f3f46;
    border-radius: 12px;
    background: #111113;
    color: #e5e7eb;
    font-family: Consolas, "Courier New", monospace;
    font-size: 12px;
}
"""

# ---------------------------
# UI widgets
# ---------------------------
class DropCard(QFrame):
    def __init__(self, on_files):
        super().__init__()
        self.setObjectName("Card")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 16)
        layout.setSpacing(10)

        title = QLabel("Convert Office files to PDF")
        title.setObjectName("Title")
        title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        hint = QLabel("Drag & drop files anywhere on this card, or choose files.")
        hint.setObjectName("Hint")

        self.choose_btn = QPushButton("Choose Files")
        self.choose_btn.setObjectName("Primary")
        self.choose_btn.clicked.connect(self.choose_files)

        header = QHBoxLayout()
        header.addWidget(title, 1)
        header.addWidget(self.choose_btn, 0, Qt.AlignRight)

        layout.addLayout(header)
        layout.addWidget(hint)

        self.on_files = on_files

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        files = [Path(u.toLocalFile()) for u in e.mimeData().urls()]
        self.on_files(files)

    def choose_files(self):
        dlg = QFileDialog(self, "Select files")
        dlg.setFileMode(QFileDialog.ExistingFiles)
        if dlg.exec():
            files = [Path(p) for p in dlg.selectedFiles()]
            self.on_files(files)

# ---------------------------
# App
# ---------------------------
class App(QWidget):
    MAX_LOG_ROWS = 10

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office → PDF (LibreOffice headless)")
        self.setMinimumSize(560, 260)        # smaller base size
        self.resize(600, 280)
        self._base_height = self.height()   # remember starting height
        self.setStyleSheet(CARD_CSS)

        # Window icon (optional; see build notes below)
        icon = self._load_icon()
        if icon:
            self.setWindowIcon(icon)

        self.soffice = find_soffice()
        if not self.soffice:
            QMessageBox.warning(
                self, "LibreOffice not found",
                "LibreOffice (soffice) was not found.\n\n"
                "Bundle it under tools/LibreOfficePortable or install it and add to PATH."
            )

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        self.drop_card = DropCard(self.convert_files)
        root.addWidget(self.drop_card)

        # Output folder row: Choose + status
        row = QHBoxLayout()
        self.choose_out_btn = QPushButton("Choose Output Folder")
        self.choose_out_btn.setObjectName("Ghost")
        self.choose_out_btn.clicked.connect(self._choose_output_dir)
        self.output_note = QLabel("Default output: source file’s folder")
        self.output_note.setObjectName("Subtle")
        row.addWidget(self.choose_out_btn, 0)
        row.addWidget(self.output_note, 1)
        root.addLayout(row)

        # Log: hidden until first item
        self.log = QListWidget()
        self.log.setVisible(False)
        self.log.setMaximumHeight(0)
        root.addWidget(self.log)

        # Footer actions
        footer = QHBoxLayout()
        self.open_btn = QPushButton("Open Output Folder")
        self.open_btn.setObjectName("Ghost")
        self.open_btn.clicked.connect(self.open_last_folder)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setObjectName("Ghost")
        self.clear_btn.clicked.connect(self._clear_log)

        footer.addStretch(1)
        footer.addWidget(self.open_btn)
        footer.addWidget(self.clear_btn)
        root.addLayout(footer)

        self._last_output_dir = None
        self._custom_output_dir = None

    # ---- icon helper
    def _load_icon(self):
        p = find_icon_path()
        return QIcon(p) if p else None

    # ---- output dir
    def _choose_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select output folder")
        if d:
            self._custom_output_dir = d
            self.output_note.setText(f"Output: {d}")
        else:
            # user cancelled; no change
            pass

    # ---- dynamic log sizing/show/hide + window growth
    def _ensure_log_visible(self):
        if not self.log.isVisible():
            self.log.setVisible(True)
        row_h = self.log.sizeHintForRow(0) if self.log.count() else 20
        rows = min(self.log.count(), self.MAX_LOG_ROWS)
        target_h = max(0, rows * row_h + 16)
        self.log.setMaximumHeight(target_h)
        # resize window height to fit content (cap at a reasonable max)
        sh = self.sizeHint()
        target_window_h = min(max(sh.height(), 260), 640)
        self.resize(self.width(), target_window_h)

    def _maybe_hide_log(self):
        if self.log.count() == 0:
            self.log.setVisible(False)
            self.log.setMaximumHeight(0)
            # snap back to original window height
            self.resize(self.width(), self._base_height)

    def _clear_log(self):
        self.log.clear()
        self._maybe_hide_log()

    # ---- actions
    def convert_files(self, files):
        if not self.soffice:
            self.log.addItem("❌ LibreOffice not found. Please install or bundle it under tools/.")
            self._ensure_log_visible()
            return
        for f in files:
            if not f.exists():
                self.log.addItem(f"❌ Missing: {f}")
                self._ensure_log_visible()
                continue
            if not is_supported(f):
                self.log.addItem(f"⏭️ Unsupported: {f.name}")
                self._ensure_log_visible()
                continue
            try:
                if self._custom_output_dir:
                    out = Path(self._custom_output_dir) / (f.stem + ".pdf")
                else:
                    out = f.with_suffix(".pdf")
                convert_with_libreoffice(f, out, self.soffice)
                self._last_output_dir = str(out.parent)
                self.log.addItem(f"✅ {f.name}  →  {out.name}")
            except subprocess.CalledProcessError:
                self.log.addItem(f"❌ {f.name}: LibreOffice failed to convert (check install).")
            except Exception as ex:
                self.log.addItem(f"❌ {f.name}: {ex}")
            finally:
                self._ensure_log_visible()

    def open_last_folder(self):
        folder = self._custom_output_dir or self._last_output_dir or os.getcwd()
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.call(["open", folder])
        else:
            subprocess.call(["xdg-open", folder])

# ---------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ico = find_icon_path()
    if ico:
        app.setWindowIcon(QIcon(ico))   # ensures window/taskbar icon
    w = App()
    w.show()
    sys.exit(app.exec())
