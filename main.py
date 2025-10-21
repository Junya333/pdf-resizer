# app_pdf_resizer_b5.py
# 依存: PyMuPDF(fitz), PySide6
import sys
import os
from pathlib import Path
import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QVBoxLayout, QHBoxLayout, QFileDialog, QComboBox, QMessageBox, QFrame
)

B5_W_MM, B5_H_MM = 182, 257
MM_TO_PT = 72 / 25.4
B5_W_PT, B5_H_PT = B5_W_MM * MM_TO_PT, B5_H_MM * MM_TO_PT


def resize_pdf_to_b5(input_path: Path, output_path: Path):
    src = fitz.open(str(input_path))
    dst = fitz.open()
    for sp in src:
        newp = dst.new_page(width=B5_W_PT, height=B5_H_PT)
        r = sp.rect
        s = min(B5_W_PT / r.width, B5_H_PT / r.height)
        dw, dh = r.width * s, r.height * s
        x0, y0 = (B5_W_PT - dw) / 2, (B5_H_PT - dh) / 2
        dest = fitz.Rect(x0, y0, x0 + dw, y0 + dh)
        newp.show_pdf_page(dest, src, sp.number, keep_proportion=True)
    dst.save(str(output_path))


class DropArea(QFrame):
    def __init__(self, on_file_dropped):
        super().__init__()
        self.on_file_dropped = on_file_dropped
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("QFrame{border:2px dashed #888;border-radius:8px;}")
        self.setAcceptDrops(True)
        lab = QLabel("「参照」を押す\nまたはここにPDFをドロップ")
        lab.setAlignment(Qt.AlignCenter)
        lay = QVBoxLayout(self)
        lay.addWidget(lab)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls() and any(u.toLocalFile().lower().endswith(".pdf") for u in e.mimeData().urls()):
            e.acceptProposedAction()

    def dropEvent(self, e: QDropEvent):
        for u in e.mimeData().urls():
            p = Path(u.toLocalFile())
            if p.suffix.lower() == ".pdf" and p.exists():
                self.on_file_dropped(p)
                break


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF B5リサイズ")
        self.setMinimumSize(QSize(520, 360))

        self.input_edit = QLineEdit()
        self.input_edit.setReadOnly(True)
        browse_btn = QPushButton("参照…")
        browse_btn.clicked.connect(self.browse_input)

        top = QHBoxLayout()
        top.addWidget(QLabel("入力PDF:"))
        top.addWidget(self.input_edit, 1)
        top.addWidget(browse_btn)

        self.drop_area = DropArea(self.set_input_file)

        self.size_combo = QComboBox()
        self.size_combo.addItem(
            "B5 (182×257mm)", userData=("B5", B5_W_PT, B5_H_PT))
        self.size_combo.currentIndexChanged.connect(
            self.refresh_default_output)

        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("(自動) 例: sample_B5.pdf")

        form = QVBoxLayout()
        form.addLayout(top)
        form.addWidget(self.drop_area)
        sz = QHBoxLayout()
        sz.addWidget(QLabel("サイズ:"))
        sz.addWidget(self.size_combo)
        form.addLayout(sz)

        out = QHBoxLayout()
        out.addWidget(QLabel("出力ファイル名:"))
        out.addWidget(self.output_name_edit, 1)
        form.addLayout(out)

        run_btn = QPushButton("出力")
        run_btn.setStyleSheet(
            "background-color:#0078d7; color:white; font-weight:bold;")
        run_btn.clicked.connect(self.run)
        form.addWidget(run_btn)

        w = QWidget()
        w.setLayout(form)
        self.setCentralWidget(w)

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "PDFを選択", str(Path.home()), "PDF (*.pdf)")
        if path:
            self.set_input_file(Path(path))

    def set_input_file(self, p: Path):
        self.input_edit.setText(str(p))
        self.refresh_default_output()

    def refresh_default_output(self):
        in_path = Path(self.input_edit.text()
                       ) if self.input_edit.text() else None
        if not in_path:
            return
        label, _, _ = self.size_combo.currentData()
        stem = in_path.stem
        default = f"{stem}_{label}.pdf"
        if not self.output_name_edit.text():
            self.output_name_edit.setText(default)

    def run(self):
        if not self.input_edit.text():
            QMessageBox.warning(self, "エラー", "入力PDFを指定してください。")
            return
        in_path = Path(self.input_edit.text())
        if not in_path.exists() or in_path.suffix.lower() != ".pdf":
            QMessageBox.critical(self, "エラー", "有効なPDFファイルを指定してください。")
            return

        label, _, _ = self.size_combo.currentData()
        out_name = self.output_name_edit.text(
        ).strip() or f"{in_path.stem}_{label}.pdf"
        # 出力は入力と同じディレクトリ
        out_path = in_path.with_name(out_name)

        if out_path.exists():
            ret = QMessageBox.question(
                self, "確認", f"{out_path.name} を上書きしますか？")
            if ret != QMessageBox.Yes:
                return

        try:
            resize_pdf_to_b5(in_path, out_path)
        except Exception as e:
            QMessageBox.critical(self, "失敗", f"出力に失敗しました。\n{e}")
            return

        QMessageBox.information(self, "完了", f"保存しました:\n{out_path}")


def main():
    # 高DPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        getattr(Qt, "HighDpiScaleFactorRoundingPolicy", None).PassThrough
        if hasattr(Qt, "HighDpiScaleFactorRoundingPolicy") else None
    )
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
