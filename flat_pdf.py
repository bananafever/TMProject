import io
import os
import re
import subprocess
import sys
from collections import Counter
import pymupdf
from PIL import Image
from PySide6 import QtCore as qtc
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw

# 렌더링 해상도. 흑백 1비트라 200 DPI도 용량 부담이 작음 (흐리면 300, 더 줄이려면 150)
DPI = 200
# 흑백 변환 임계값 (0~255). 이보다 밝으면 흰색. 얇은 선/회색 글자가 사라지면 값을 올림
THRESHOLD = 180
_BW_LUT = [255 if v > THRESHOLD else 0 for v in range(256)]

# 예상 용량 계산 시 실제로 변환해 볼 최대 페이지 수 (문서 전체에 고르게 샘플링)
SAMPLE_PAGES = 5
# 이미지가 페이지 면적의 이 비율 이상을 덮으면 스캔 페이지로 간주
SCAN_COVERAGE = 0.5

PAGE_LABELS = {
    "text": "텍스트",
    "ocr": "OCR",
    "image": "이미지",
    "other": "빈/도형",
}
DOC_LABELS = {
    "text": "텍스트 PDF",
    "ocr": "OCR된 PDF",
    "image": "OCR 안 된 이미지 PDF",
    "other": "텍스트 없는 도형/빈 PDF",
}


def _render_bw(page):
    # RGB 대신 회색조로 렌더링한 뒤 임계값으로 1비트 흑백 변환
    pix = page.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY, alpha=False)
    gray = Image.frombuffer("L", (pix.width, pix.height), pix.samples, "raw", "L", pix.stride, 1)
    return gray.point(_BW_LUT, mode="1")


def _save_bw_pdf(images, dst):
    # Pillow는 1비트 이미지를 CCITT Group 4(팩스 압축)로 PDF에 기록함
    # resolution=DPI로 원본 페이지의 물리 크기 유지
    images[0].save(dst, "PDF", save_all=True, append_images=images[1:], resolution=DPI)


def flatten_pdf(src_path, dst_path):
    """각 페이지를 1비트 흑백 이미지로 바꿔 CCITT G4 압축 PDF로 저장 (텍스트 검색 불가)."""
    with pymupdf.open(src_path) as doc:
        images = [_render_bw(page) for page in doc]
    _save_bw_pdf(images, dst_path)


def _color_label(doc, img):
    """페이지 이미지의 색 형식 (흑백/회색조/컬러)."""
    xref, bpc, cs_name = img[0], img[4], img[5]
    if bpc == 1:
        return "흑백"
    if "Gray" in cs_name:
        return "회색조"
    if cs_name == "ICCBased":
        # ICC 프로파일의 채널 수(N)로 회색조/컬러 구분: [/ICCBased 12 0 R]
        kind, cs = doc.xref_get_key(xref, "ColorSpace")
        if kind == "xref":  # 색공간 배열이 별도 객체로 분리된 경우 (예: 7 0 R)
            cs = doc.xref_object(int(cs.split()[0]))
        m = re.search(r"/ICCBased\s+(\d+)\s+0\s+R", cs)
        if m and doc.xref_get_key(int(m.group(1)), "N")[1] == "1":
            return "회색조"
    return "컬러"


def _classify_page(doc, page):
    """페이지 유형과 (스캔 페이지라면) 이미지 색 형식·해상도를 반환."""
    has_text = bool(page.get_text().strip())
    page_area = abs(page.rect)
    for img in page.get_images(full=True):
        for rect in page.get_image_rects(img[0]):
            if abs(rect & page.rect) >= page_area * SCAN_COVERAGE:
                # 이미지 픽셀 수 ÷ 표시 크기(인치)로 실제 해상도 계산
                dpi = max(img[2], img[3]) / (max(rect.width, rect.height) / 72)
                return ("ocr" if has_text else "image"), (_color_label(doc, img), dpi)
    return ("text" if has_text else "other"), None


def analyze_pdf(path):
    """PDF 유형을 판별하고, 일부 페이지를 실제 변환해 변환 후 용량을 추정."""
    with pymupdf.open(path) as doc:
        n = len(doc)
        if n == 0:
            raise ValueError("페이지가 없는 PDF입니다.")

        kinds = Counter()
        scans = []
        for page in doc:
            kind, scan = _classify_page(doc, page)
            kinds[kind] += 1
            if scan:
                scans.append(scan)

        k = min(SAMPLE_PAGES, n)
        sample = sorted({round(i * (n - 1) / (k - 1)) for i in range(k)}) if k > 1 else [0]
        buf = io.BytesIO()
        _save_bw_pdf([_render_bw(doc[i]) for i in sample], buf)
        est_size = len(buf.getvalue()) / len(sample) * n

    return {
        "pages": n,
        "kinds": kinds,
        "scans": scans,
        "orig_size": os.path.getsize(path),
        "est_size": est_size,
    }


def summarize(info):
    """analyze_pdf 결과를 화면에 표시할 항목별 문구로 정리."""
    kinds = info["kinds"]
    if len(kinds) == 1:
        doc_type = DOC_LABELS[next(iter(kinds))]
    else:
        detail = ", ".join(f"{PAGE_LABELS[k]} {c}쪽" for k, c in kinds.most_common())
        doc_type = f"혼합 ({detail})"

    scan = "없음"
    if info["scans"]:
        colors = Counter(c for c, _ in info["scans"])
        avg_dpi = sum(d for _, d in info["scans"]) / len(info["scans"])
        scan = f"{', '.join(colors)}, 약 {avg_dpi:.0f} DPI"

    orig, est = info["orig_size"], info["est_size"]
    ratio = est / orig if orig else 1
    if ratio < 0.9:
        level, verdict = "ok", "크기를 줄일 수 있습니다"
    elif ratio <= 1.0:
        level, verdict = "warn", "크기가 거의 줄지 않습니다"
    else:
        level, verdict = "bad", "원본보다 커집니다"

    notes = []
    if kinds["text"]:
        notes.append("텍스트 페이지는 이미지로 바꾸면 용량이 커집니다.")
    if kinds["ocr"]:
        notes.append("OCR 페이지는 변환 시 숨은 텍스트 층이 제거됩니다.")
    if kinds["image"] and len(kinds) == 1:
        hint = " 더 줄이려면 DPI 값을 낮추세요." if ratio >= 0.9 else ""
        notes.append(f"이미 검색 불가 PDF입니다.{hint}")

    return {
        "pages": info["pages"],
        "doc_type": doc_type,
        "scan": scan,
        "orig_mb": orig / (1024 * 1024),
        "est_mb": est / (1024 * 1024),
        "pct": (ratio - 1) * 100,
        "level": level,
        "verdict": verdict,
        "notes": notes,
    }


def convert_pdf(src_path):
    """원본 파일명 뒤에 _flat을 붙여 변환 저장하고 (저장 경로, 용량)을 반환."""
    base_path, ext = os.path.splitext(src_path)
    output_path = os.path.normpath(f"{base_path}_flat{ext}")
    flatten_pdf(src_path, output_path)
    return output_path, os.path.getsize(output_path)


# ---------------------------------------------------------------------------
# 화면 (PySide6)
# ---------------------------------------------------------------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def setup_app_fonts(app):
    """assets/fonts에 번들된 Pretendard를 등록하고 앱 전역 글꼴로 지정합니다."""
    font_dir = resource_path(os.path.join("assets", "fonts"))
    if os.path.isdir(font_dir):
        for f in os.listdir(font_dir):
            if f.lower().endswith((".ttf", ".otf")):
                qtg.QFontDatabase.addApplicationFont(os.path.join(font_dir, f))

    font = app.font()
    font.setFamily("Pretendard")
    font.setStyleStrategy(qtg.QFont.PreferAntialias)
    font.setHintingPreference(qtg.QFont.PreferNoHinting)
    app.setFont(font)


# main.py의 Clean Light 팔레트 (웜 뉴트럴 캔버스 + 틸 액센트)
STYLE_SHEET = """
    QWidget {
        background-color: #F6F6F4;
        color: #333333;
        font-family: 'Pretendard', 'Malgun Gothic', sans-serif;
        font-size: 10pt;
    }
    #card {
        background-color: white;
        border: 1px solid #DCE3DF;
        border-radius: 12px;
    }
    #card QLabel { background: transparent; }
    #title { color: #232220; font-size: 15pt; font-weight: bold; }
    #subtitle, #muted { color: #8A877E; font-size: 9pt; }
    #fieldLabel { color: #8A877E; }
    #fileName { color: #232220; font-weight: bold; }
    #dropZone {
        background-color: white;
        border: 2px dashed #C9CFCB;
        border-radius: 12px;
    }
    #dropZone[active="true"] {
        background-color: #E1EEE7;
        border: 2px dashed #26695E;
    }
    #dropZone QLabel { background: transparent; }
    #dropIcon { color: #26695E; font-size: 22pt; }
    #banner { border-radius: 10px; padding: 9px 12px; }
    #banner[level="ok"] { background-color: #E1EEE7; color: #1B4B43; }
    #banner[level="warn"] { background-color: #FAEEDA; color: #854F0B; }
    #banner[level="bad"] { background-color: #FAECE7; color: #993C1D; }
    #banner[level="busy"] { background-color: #ECEAE5; color: #4B4A47; }
    QPushButton {
        background-color: #26695E;
        color: white;
        border: none;
        border-radius: 10px;
        padding: 8px 18px;
        font-weight: bold;
    }
    QPushButton:hover { background-color: #2F7C6E; }
    QPushButton:pressed { background-color: #1B4B43; }
    QPushButton:disabled { background-color: #B4C9C3; }
    QPushButton#secondary {
        background-color: white;
        color: #333333;
        border: 1px solid #DCE3DF;
        font-weight: normal;
        padding: 5px 12px;
    }
    QPushButton#secondary:hover { background-color: #ECEAE5; }
    QPushButton#secondary:disabled { color: #B4B2A9; }
    QProgressBar {
        background-color: #ECEAE5;
        border: none;
        border-radius: 3px;
        max-height: 6px;
    }
    QProgressBar::chunk { background-color: #26695E; border-radius: 3px; }
"""


def _label(text="", name=None, wrap=False):
    lbl = qtw.QLabel(text)
    if name:
        lbl.setObjectName(name)
    lbl.setWordWrap(wrap)
    return lbl


def _set_prop(widget, name, value):
    """스타일시트의 [속성] 선택자가 다시 적용되도록 속성 변경 후 재polish."""
    widget.setProperty(name, value)
    widget.style().unpolish(widget)
    widget.style().polish(widget)


class Worker(qtc.QObject):
    """오래 걸리는 판별/변환 작업을 별도 스레드에서 실행."""
    finished = qtc.Signal(object)
    failed = qtc.Signal(str)

    def __init__(self, fn, *args):
        super().__init__()
        self._fn = fn
        self._args = args

    @qtc.Slot()
    def run(self):
        try:
            self.finished.emit(self._fn(*self._args))
        except Exception as e:
            self.failed.emit(str(e))


class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 용량 줄이기")
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(qtg.QIcon(icon_path))
        self.setFixedWidth(460)
        self.setAcceptDrops(True)

        self._path = None
        self._summary = None
        self._output = None
        self._job = None

        self.stack = qtw.QStackedWidget()
        self.stack.addWidget(self._build_start_page())
        self.stack.addWidget(self._build_result_page())
        layout = qtw.QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.addWidget(self.stack)

    # ----- 화면 구성 -----
    def _build_start_page(self):
        page = qtw.QWidget()
        v = qtw.QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)
        v.addWidget(_label("PDF 용량 줄이기", "title"))
        v.addWidget(_label("검색되지 않는 흑백 이미지 PDF로 변환합니다", "subtitle"))
        v.addSpacing(16)

        self.drop_zone = qtw.QFrame()
        self.drop_zone.setObjectName("dropZone")
        dz = qtw.QVBoxLayout(self.drop_zone)
        dz.setContentsMargins(16, 32, 16, 32)
        dz.setSpacing(6)
        for w in (
            _label("⇪", "dropIcon"),
            _label("PDF 파일을 여기에 끌어다 놓거나"),
            _label("선택하면 바로 유형과 예상 용량을 판별합니다", "muted"),
        ):
            w.setAlignment(qtc.Qt.AlignCenter)
            dz.addWidget(w)
        dz.addSpacing(10)
        pick_btn = qtw.QPushButton("PDF 파일 선택")
        pick_btn.setCursor(qtc.Qt.PointingHandCursor)
        pick_btn.clicked.connect(self.pick_file)
        dz.addWidget(pick_btn, alignment=qtc.Qt.AlignCenter)
        v.addWidget(self.drop_zone)

        v.addSpacing(12)
        v.addWidget(_label(f"변환 설정: {DPI} DPI · 흑백 1비트 (CCITT G4)", "muted"))
        v.addStretch()
        return page

    def _build_result_page(self):
        page = qtw.QWidget()
        v = qtw.QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        # 파일 카드
        card = qtw.QFrame()
        card.setObjectName("card")
        h = qtw.QHBoxLayout(card)
        h.setContentsMargins(14, 10, 10, 10)
        info = qtw.QVBoxLayout()
        info.setSpacing(2)
        self.file_name = _label(name="fileName")
        self.file_meta = _label(name="muted")
        info.addWidget(self.file_name)
        info.addWidget(self.file_meta)
        h.addLayout(info, 1)
        self.other_btn = qtw.QPushButton("다른 파일")
        self.other_btn.setObjectName("secondary")
        self.other_btn.clicked.connect(self.reset)
        h.addWidget(self.other_btn, alignment=qtc.Qt.AlignVCenter)
        v.addWidget(card)

        # 판별 결과
        form = qtw.QFormLayout()
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(8)
        self.val_type = _label(wrap=True)
        self.val_scan = _label(wrap=True)
        self.val_size = _label(wrap=True)
        for text, val in (("문서 유형", self.val_type), ("스캔 이미지", self.val_scan), ("예상 용량", self.val_size)):
            form.addRow(_label(text, "fieldLabel"), val)
        v.addLayout(form)

        self.banner = _label(name="banner", wrap=True)
        v.addWidget(self.banner)
        self.notes = _label(name="muted", wrap=True)
        self.notes.setTextInteractionFlags(qtc.Qt.TextSelectableByMouse)
        v.addWidget(self.notes)

        self.progress = qtw.QProgressBar()
        self.progress.setRange(0, 0)  # 진행률을 알 수 없으므로 무한 진행 표시
        self.progress.setTextVisible(False)
        v.addWidget(self.progress)

        v.addStretch()
        buttons = qtw.QHBoxLayout()
        self.folder_btn = qtw.QPushButton("폴더 열기")
        self.folder_btn.setObjectName("secondary")
        self.folder_btn.clicked.connect(self.open_folder)
        buttons.addWidget(self.folder_btn)
        buttons.addStretch()
        self.convert_btn = qtw.QPushButton("변환하기")
        self.convert_btn.setCursor(qtc.Qt.PointingHandCursor)
        self.convert_btn.clicked.connect(self.convert)
        buttons.addWidget(self.convert_btn)
        v.addLayout(buttons)
        return page

    # ----- 파일 입력 -----
    def pick_file(self):
        path, _ = qtw.QFileDialog.getOpenFileName(
            self, "변환할 PDF 파일을 선택하세요", "", "PDF 파일 (*.pdf);;모든 파일 (*.*)"
        )
        if path:
            self.load_file(path)

    def _dropped_pdf(self, event):
        urls = event.mimeData().urls()
        if len(urls) == 1 and urls[0].isLocalFile():
            path = urls[0].toLocalFile()
            if path.lower().endswith(".pdf"):
                return path
        return None

    def dragEnterEvent(self, event):
        if self._job is None and self._dropped_pdf(event):
            event.acceptProposedAction()
            _set_prop(self.drop_zone, "active", True)

    def dragLeaveEvent(self, event):
        _set_prop(self.drop_zone, "active", False)

    def dropEvent(self, event):
        _set_prop(self.drop_zone, "active", False)
        path = self._dropped_pdf(event)
        if path:
            self.load_file(path)

    # ----- 상태 전환 -----
    def reset(self):
        self._path = self._summary = self._output = None
        self.stack.setCurrentIndex(0)

    def load_file(self, path):
        self._path = path
        self._summary = self._output = None
        name = os.path.basename(path)
        self.file_name.setText(self.file_name.fontMetrics().elidedText(name, qtc.Qt.ElideMiddle, 300))
        self.file_name.setToolTip(path)
        self.file_meta.setText(f"{os.path.getsize(path) / (1024 * 1024):.2f} MB")
        for val in (self.val_type, self.val_scan, self.val_size):
            val.setText("—")
        self._set_banner("busy", "판별 중…")
        self.notes.clear()
        self.folder_btn.hide()
        self.convert_btn.show()
        self.stack.setCurrentIndex(1)
        self._run(analyze_pdf, path, on_done=self._on_analyzed)

    def _on_analyzed(self, info):
        s = self._summary = summarize(info)
        self.file_meta.setText(f"{s['pages']}쪽 · {s['orig_mb']:.2f} MB")
        self.val_type.setText(s["doc_type"])
        self.val_scan.setText(s["scan"])
        self.val_size.setText(f"{s['orig_mb']:.2f} MB → {s['est_mb']:.2f} MB ({s['pct']:+.0f}%)")
        mark = {"ok": "✓", "warn": "△", "bad": "✗"}[s["level"]]
        self._set_banner(s["level"], f"{mark}  {s['verdict']}")
        self.notes.setText("\n".join(f"※ {n}" for n in s["notes"]))

    def convert(self):
        self._set_banner("busy", "변환 중…")
        self._run(convert_pdf, self._path, on_done=self._on_converted)

    def _on_converted(self, result):
        self._output, size = result
        new_mb = size / (1024 * 1024)
        pct = (new_mb / self._summary["orig_mb"] - 1) * 100 if self._summary["orig_mb"] else 0
        self._set_banner("ok", f"✓  변환 완료 · {new_mb:.2f} MB ({pct:+.0f}%)")
        self.notes.setText(f"저장 위치: {self._output}")
        self.convert_btn.hide()
        self.folder_btn.show()

    def _on_failed(self, message):
        self._set_banner("bad", f"✗  처리 중 오류가 발생했습니다: {message}")

    def open_folder(self):
        if self._output:
            # 탐색기에서 결과 파일을 선택한 상태로 폴더 열기
            subprocess.Popen(["explorer", "/select,", os.path.normpath(self._output)])

    def _set_banner(self, level, text):
        self.banner.setText(text)
        _set_prop(self.banner, "level", level)

    # ----- 백그라운드 작업 -----
    def _run(self, fn, *args, on_done):
        thread = qtc.QThread(self)
        worker = Worker(fn, *args)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(on_done)
        worker.failed.connect(self._on_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._on_job_end)
        self._job = (thread, worker)
        self._set_busy(True)
        thread.start()

    def _on_job_end(self):
        self._job = None
        self._set_busy(False)

    def _set_busy(self, busy):
        self.progress.setVisible(busy)
        self.other_btn.setEnabled(not busy)
        # 판별이 끝난 뒤에만 변환 가능
        self.convert_btn.setEnabled(not busy and self._summary is not None)

    def closeEvent(self, event):
        # 작업 중 창을 닫으면 스레드가 끝날 때까지 기다린 뒤 종료
        if self._job:
            self._job[0].wait()
        super().closeEvent(event)


def main():
    app = qtw.QApplication(sys.argv)
    setup_app_fonts(app)
    app.setStyleSheet(STYLE_SHEET)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
