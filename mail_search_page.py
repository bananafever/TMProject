"""Search 탭 — Microsoft Graph `$search` 로 메일함 전체나 한 폴더(하위 폴더까지)를 찾고, 앱 안에서 미리 본다(10-25).

Outlook 검색창이 불편해서 만들었다. 읽기만 한다 — 메일을 옮기지·지우지 않는다. 찾기·본문 받기·저장은 스레드에서,
화면 고치기는 GUI 스레드에서(Signal). 메일함은 main 이 넘겨 준 `get_mailbox()` 로 받는다 — 클래식 Outlook 으로 읽는
중이어도 Microsoft 로그인이 남아 있으면 쓸 수 있다.
"""

import logging
import re
import threading
import time
from datetime import date
from html import escape
from pathlib import Path

from dateutil.relativedelta import relativedelta
from PySide6 import QtCore as qtc
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw

import graph_mail

PERIODS = [("전체 기간", None), ("최근 1개월", relativedelta(months=1)), ("최근 6개월", relativedelta(months=6)),
           ("최근 1년", relativedelta(years=1)), ("최근 3년", relativedelta(years=3))]
COLUMNS = ["받은 날", "보낸 사람", "제목", "폴더", "첨부"]
BAD_NAME = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def safe_name(name: str, fallback="mail") -> str:
    """파일 이름에 못 쓰는 글자를 지운다 — 제목이 그대로 .eml 이름이 된다."""
    name = BAD_NAME.sub(" ", name or "").strip(" .")
    return name[:120] or fallback


def free_path(folder: Path, name: str) -> Path:
    """같은 이름이 있으면 「이름 (2).pdf」 처럼 — 저장한 파일을 덮지 않는다."""
    path = folder / name
    n = 2
    while path.exists():
        path = folder / f"{Path(name).stem} ({n}){Path(name).suffix}"
        n += 1
    return path


def format_mail_date(dt) -> str:
    """수신 일시를 컴팩트하게 포맷한다 — 오늘이면 시간만, 올해면 월-일 시:분, 이전 해는 연-월-일."""
    if not dt:
        return ""
    today = date.today()
    if dt.date() == today:
        return dt.strftime("%H:%M")
    if dt.year == today.year:
        return dt.strftime("%m-%d %H:%M")
    return dt.strftime("%Y-%m-%d")


def size_text(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / 1024 / 1024:.1f}MB"
    return f"{max(1, round(size / 1024))}KB"


class _BorderOverlay(qtw.QWidget):
    """RoundedFrame 위에 얹는 테두리 — 마우스는 통과시키고 안티에일리어싱으로 한 번만 그린다."""

    def __init__(self, frame):
        super().__init__(frame)
        self.frame = frame
        self.setAttribute(qtc.Qt.WA_TransparentForMouseEvents)
        self.setAttribute(qtc.Qt.WA_NoSystemBackground)

    def paintEvent(self, _event):
        p = qtg.QPainter(self)
        p.setRenderHint(qtg.QPainter.Antialiasing)
        pen = qtg.QPen(qtg.QColor(self.frame.border_color), self.frame.border_width)
        p.setPen(pen)
        p.setBrush(qtc.Qt.NoBrush)
        half = self.frame.border_width / 2
        rect = qtc.QRectF(self.rect()).adjusted(half, half, -half, -half)
        p.drawRoundedRect(rect, self.frame.radius - half, self.frame.radius - half)


class RoundedFrame(qtw.QFrame):
    """자식 위젯을 둥근 모양으로 잘라 넣고 균일한 테두리를 위에 그린다.
    QTableWidget 에 border-radius 를 주면 헤더·행이 모서리를 덮어 테두리가 끊기거나 두꺼워 보인다(10-28)."""

    def __init__(self, radius=12, border_color="#DCE3DF", border_width=1, parent=None):
        super().__init__(parent)
        self.radius, self.border_color, self.border_width = radius, border_color, border_width
        self._child = None
        layout = qtw.QVBoxLayout(self)
        layout.setContentsMargins(border_width, border_width, border_width, border_width)
        layout.setSpacing(0)
        self._overlay = _BorderOverlay(self)

    def set_widget(self, widget):
        self._child = widget
        self.layout().addWidget(widget)
        self._overlay.raise_()

    def paintEvent(self, _event):
        p = qtg.QPainter(self)
        p.setRenderHint(qtg.QPainter.Antialiasing)
        p.setPen(qtc.Qt.NoPen)
        p.setBrush(qtg.QColor("white"))
        p.drawRoundedRect(qtc.QRectF(self.rect()), self.radius, self.radius)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._overlay.setGeometry(self.rect())
        self._overlay.raise_()
        if self._child is not None:
            # 자식은 테두리 안쪽 둥근 모양으로 자른다 — 거친 가장자리는 위의 테두리가 덮는다
            r = self.radius - self.border_width
            path = qtg.QPainterPath()
            path.addRoundedRect(qtc.QRectF(self._child.rect()), r, r)
            self._child.setMask(qtg.QRegion(path.toFillPolygon().toPolygon()))


class MailSearchPage(qtw.QWidget):
    folders_ready = qtc.Signal(list)
    found = qtc.Signal(object, float)
    failed = qtc.Signal(str)
    preview_ready = qtc.Signal(object, str, object)
    saved = qtc.Signal(str)

    def __init__(self, get_mailbox, parent=None):
        super().__init__(parent)
        self.get_mailbox = get_mailbox
        self._busy = False
        self._folders_loaded = False
        self._current = None
        self._attachments = None
        self.folders_ready.connect(self._fill_folders)
        self.found.connect(self._show_results)
        self.failed.connect(self._show_error)
        self.preview_ready.connect(self._show_preview)
        self.saved.connect(self._show_saved)
        self._build()

    # ── 화면 ──

    def _build(self):
        layout = qtw.QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        # 상단 통합 검색 및 필터 바
        filter_box = qtw.QFrame()
        filter_box.setObjectName("searchFilterContainer")
        filter_layout = qtw.QVBoxLayout(filter_box)
        filter_layout.setContentsMargins(14, 12, 14, 12)
        filter_layout.setSpacing(10)

        # 1행: 검색어 + 폴더 콤보 + 새로고침 + 검색 버튼
        row1 = qtw.QHBoxLayout()
        row1.setSpacing(8)
        self.query = qtw.QLineEdit()
        self.query.setPlaceholderText("검색어 입력 — from:이름  subject:제목  body:본문")
        self.query.setClearButtonEnabled(True)
        self.query.returnPressed.connect(self.search)

        self.folder_combo = qtw.QComboBox()
        self.folder_combo.setView(qtw.QListView())
        self.folder_combo.setMinimumWidth(200)
        self.folder_combo.addItem("전체 메일함", None)
        self.folder_combo.currentIndexChanged.connect(self._folder_changed)

        self.refresh_btn = qtw.QPushButton("↻")
        self.refresh_btn.setObjectName("secondaryButton")
        self.refresh_btn.setToolTip("폴더 목록 새로 받기")
        self.refresh_btn.setFixedWidth(34)
        self.refresh_btn.clicked.connect(lambda: self.load_folders(refresh=True))

        self.search_btn = qtw.QPushButton("검색")
        self.search_btn.setFixedWidth(72)
        self.search_btn.clicked.connect(self.search)

        row1.addWidget(self.query, 1)
        row1.addWidget(self.folder_combo)
        row1.addWidget(self.refresh_btn)
        row1.addWidget(self.search_btn)
        filter_layout.addLayout(row1)

        # 2행: 세부 필터 (기간, 첨부 여부, 하위 폴더 포함, 보낸 사람) + 상태 라벨
        row2 = qtw.QHBoxLayout()
        row2.setSpacing(12)

        period_lbl = qtw.QLabel("기간:")
        period_lbl.setStyleSheet("color: #64748B; font-weight: 500;")
        self.period = qtw.QComboBox()
        self.period.setView(qtw.QListView())
        self.period.setMinimumWidth(110)
        for label, delta in PERIODS:
            self.period.addItem(label, delta)
        self.period.setCurrentIndex(3)

        self.has_attachments = qtw.QCheckBox("📎 첨부 있음")
        self.subfolders = qtw.QCheckBox("하위 폴더 포함")
        self.subfolders.setChecked(True)
        self.subfolders.setEnabled(False)

        self.sender = qtw.QLineEdit()
        self.sender.setPlaceholderText("보낸 사람")
        self.sender.setClearButtonEnabled(True)
        self.sender.setFixedWidth(140)
        self.sender.returnPressed.connect(self.search)

        self.status = qtw.QLabel("")
        self.status.setObjectName("searchStatus")

        row2.addWidget(period_lbl)
        row2.addWidget(self.period)
        row2.addWidget(self.has_attachments)
        row2.addWidget(self.subfolders)
        row2.addWidget(self.sender)
        row2.addStretch()
        row2.addWidget(self.status)
        filter_layout.addLayout(row2)

        layout.addWidget(filter_box)

        # ── 메인 바디: 좌우 2단 분할 (Horizontal Splitter) ──
        splitter = qtw.QSplitter(qtc.Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)

        # 좌측: 메일 목록 — RoundedFrame 이 둥글게 자르고 테두리를 한 번만 그린다
        self.table_frame = RoundedFrame(radius=12)
        self.table_frame.setObjectName("tableContainer")

        self.table = qtw.QTableWidget(0, len(COLUMNS))
        self.table.setObjectName("mailTable")
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setFrameShape(qtw.QFrame.NoFrame)
        header = self.table.horizontalHeader()
        font = qtw.QApplication.font()
        self.table.setFont(font)
        header.setFont(font)

        header.setMinimumSectionSize(20)
        header.setSectionResizeMode(2, qtw.QHeaderView.Stretch)
        for col, width in ((0, 95), (1, 85), (3, 65), (4, 38)):
            header.setSectionResizeMode(col, qtw.QHeaderView.Interactive)
            header.resizeSection(col, width)

        header.setHighlightSections(False)
        header.setSortIndicator(0, qtc.Qt.DescendingOrder)
        header.setDefaultAlignment(qtc.Qt.AlignLeft | qtc.Qt.AlignVCenter)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setHorizontalScrollBarPolicy(qtc.Qt.ScrollBarAlwaysOff)
        self.table.setSelectionBehavior(qtw.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(qtw.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(qtw.QAbstractItemView.NoEditTriggers)
        self.table.currentCellChanged.connect(lambda row, *_: self._select(row))
        self.table_frame.set_widget(self.table)
        splitter.addWidget(self.table_frame)

        # 우측: 메일 상세 미리보기 패널
        self.preview = qtw.QFrame()
        self.preview.setObjectName("previewContainer")
        box = qtw.QVBoxLayout(self.preview)
        box.setContentsMargins(16, 14, 16, 14)
        box.setSpacing(8)

        # 미리보기 헤더: 제목 + 액션 버튼
        top = qtw.QHBoxLayout()
        top.setSpacing(10)
        self.pv_subject = qtw.QLabel("메일을 선택하면 여기에 표시됩니다")
        self.pv_subject.setObjectName("previewSubject")
        self.pv_subject.setTextInteractionFlags(qtc.Qt.TextSelectableByMouse)
        self.pv_subject.setWordWrap(True)

        self.save_attach_btn = qtw.QPushButton("첨부 저장")
        self.save_attach_btn.setObjectName("secondaryButton")
        self.save_attach_btn.clicked.connect(self.save_attachments)

        self.save_eml_btn = qtw.QPushButton(".eml 저장")
        self.save_eml_btn.setObjectName("secondaryButton")
        self.save_eml_btn.clicked.connect(self.save_eml)

        top.addWidget(self.pv_subject, 1)
        top.addWidget(self.save_attach_btn)
        top.addWidget(self.save_eml_btn)
        box.addLayout(top)

        # 메타 정보 (보낸이, 수신일시, 폴더 등)
        self.pv_meta = qtw.QLabel("")
        self.pv_meta.setObjectName("previewMeta")
        self.pv_meta.setWordWrap(True)
        self.pv_meta.setTextInteractionFlags(qtc.Qt.TextSelectableByMouse)
        box.addWidget(self.pv_meta)

        # 첨부파일 요약 칩
        self.pv_files = qtw.QLabel("")
        self.pv_files.setObjectName("previewFiles")
        self.pv_files.setWordWrap(True)
        self.pv_files.setVisible(False)
        box.addWidget(self.pv_files)

        # 메일 본문 브라우저
        self.pv_body = qtw.QTextBrowser()
        self.pv_body.setObjectName("previewBody")
        self.pv_body.setOpenExternalLinks(True)
        self.pv_body.setOpenLinks(True)
        box.addWidget(self.pv_body, 1)

        self._set_buttons(False)
        self._clear_preview()

        splitter.addWidget(self.preview)
        splitter.setSizes([380, 360])
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        layout.addWidget(splitter, 1)

    def _set_buttons(self, on):
        self.save_eml_btn.setEnabled(on)
        self.save_attach_btn.setEnabled(on and bool(self._attachments))

    def focus_query(self):
        self.query.setFocus()
        self.query.selectAll()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._folders_loaded:
            self._folders_loaded = True
            self.load_folders()

    def _folder_changed(self, _index):
        self.subfolders.setEnabled(self.folder_combo.currentData() is not None)

    # ── 폴더 ──

    def load_folders(self, refresh=False):
        def run():
            try:
                box = self.get_mailbox()
                if not box:
                    raise graph_mail.LoginRequired("메일함이 연결되지 않았습니다.")
                self.folders_ready.emit(box.folders(refresh=refresh))
            except Exception as e:                       # noqa: BLE001
                logging.error(f"메일 폴더 목록을 받지 못했습니다: {e}")
                self.failed.emit(self._error_text(e))
        threading.Thread(target=run, daemon=True).start()

    def _fill_folders(self, folders):
        keep = self.folder_combo.currentData()
        self.folder_combo.blockSignals(True)
        self.folder_combo.clear()
        self.folder_combo.addItem("전체 메일함", None)
        for f in folders:
            self.folder_combo.addItem("    " * (f["depth"] + 1) + f["name"], f["id"])
            self.folder_combo.setItemData(self.folder_combo.count() - 1, f["path"], qtc.Qt.ToolTipRole)
        self.folder_combo.setCurrentIndex(max(0, self.folder_combo.findData(keep)))
        self.folder_combo.blockSignals(False)
        self._folder_changed(self.folder_combo.currentIndex())

    # ── 검색 ──

    def search(self):
        if self._busy:
            return
        delta = self.period.currentData()
        args = dict(text=self.query.text(), folder_id=self.folder_combo.currentData(),
                    subfolders=self.subfolders.isChecked(), since=(date.today() - delta) if delta else None,
                    has_attachments=self.has_attachments.isChecked(), sender=self.sender.text())
        if not graph_mail.kql(args["text"], args["sender"], args["since"], args["has_attachments"]):
            self._show_error("검색어나 보낸 사람·기간·첨부 조건을 넣으세요.")
            return
        self._busy = True
        self.search_btn.setEnabled(False)
        self.status.setText("찾는 중…")

        def run():
            started = time.monotonic()
            try:
                box = self.get_mailbox()
                if not box:
                    raise graph_mail.LoginRequired("메일함이 연결되지 않았습니다.")
                result = box.search(**args)
                box.folders()                            # 결과의 폴더 이름을 보이려고 — 이미 받았으면 그대로
                self.found.emit(result, time.monotonic() - started)
            except Exception as e:                       # noqa: BLE001
                logging.error(f"메일 검색 실패: {e}")
                self.failed.emit(self._error_text(e))
        threading.Thread(target=run, daemon=True).start()

    @staticmethod
    def _error_text(e):
        if isinstance(e, graph_mail.LoginRequired):
            return "Microsoft 로그인이 필요합니다 — Settings 「메일」 에서 로그인하세요."
        return f"찾지 못했습니다: {e}"

    def _show_error(self, text):
        self._busy = False
        self.search_btn.setEnabled(True)
        self.status.setText(text)

    def _show_results(self, result, seconds):
        self._busy = False
        self.search_btn.setEnabled(True)
        count = f"{len(result)}통 이상 — 조건을 좁혀 보세요" if result.truncated else f"{len(result)}통"
        self.status.setText(f"{count} · {seconds:.1f}초")
        box = self.get_mailbox()
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(len(result))
        for r, message in enumerate(result):
            path = box.folder_path(message.parent_folder_id)
            received = message.ReceivedTime
            cells = [format_mail_date(received), message.SenderName, message.Subject,
                     path.rsplit(" / ", 1)[-1], "📎" if message.data.get("hasAttachments") else ""]
            for c, text in enumerate(cells):
                item = qtw.QTableWidgetItem(text)
                if c == 0:
                    item.setData(qtc.Qt.UserRole, message)
                    if received:
                        item.setToolTip(f"수신일시: {received.strftime('%Y-%m-%d %H:%M:%S')}")
                elif c == 1:
                    item.setToolTip(message.SenderName)
                elif c == 2:
                    item.setToolTip(message.Preview or message.Subject)
                elif c == 3:
                    item.setToolTip(path)
                elif c == 4:
                    item.setTextAlignment(qtc.Qt.AlignCenter)
                    if text:
                        item.setToolTip("첨부파일 있음")
                self.table.setItem(r, c, item)
        self.table.setSortingEnabled(True)
        self.table.scrollToTop()
        if result:
            self.table.selectRow(0)
        else:
            self._clear_preview("찾은 메일이 없습니다")

    # ── 미리보기 ──

    def _message_at(self, row):
        item = self.table.item(row, 0) if row >= 0 else None
        return item.data(qtc.Qt.UserRole) if item is not None else None

    def _clear_preview(self, title="메일을 선택하면 여기에 표시됩니다"):
        self._current, self._attachments = None, None
        self.pv_subject.setText(title)
        self.pv_meta.setText("")
        self.pv_files.setText("")
        self.pv_files.setVisible(False)
        self.pv_body.setHtml(
            "<div style='text-align: center; color: #94A3B8; margin-top: 60px; font-family: sans-serif;'>"
            "<p style='font-size: 32px; margin-bottom: 6px;'>✉️</p>"
            "<p style='font-weight: 600; color: #64748B; font-size: 13px; margin-bottom: 4px;'>선택된 메일이 없습니다</p>"
            "<p style='font-size: 11px; color: #94A3B8;'>목록에서 메일을 클릭하여 본문과 첨부파일을 확인하세요.</p>"
            "</div>"
        )
        self._set_buttons(False)

    def _select(self, row):
        message = self._message_at(row)
        if message is None:
            return
        if message is self._current:
            return
        self._current, self._attachments = message, None
        stamp = message.ReceivedTime
        received = f"{stamp:%Y-%m-%d} ({'월화수목금토일'[stamp.weekday()]}) {stamp:%H:%M}" if stamp else ""
        path = self.get_mailbox().folder_path(message.parent_folder_id)
        meta = []
        if message.SenderName:
            meta.append(f"<b>보낸 사람:</b> {escape(message.SenderName)}")
        if received:
            meta.append(f"<b>수신:</b> {received}")
        if path:
            folder_name = path.rsplit(" / ", 1)[-1]
            meta.append(f"<b>폴더:</b> {escape(folder_name)}")
        if message.To:
            meta.append(f"<b>받는 사람:</b> {escape(message.To)}")
        self.pv_subject.setText(message.Subject or "(제목 없음)")
        self.pv_meta.setText(" &nbsp;·&nbsp; ".join(meta))
        if message.data.get("hasAttachments"):
            self.pv_files.setText("📎 첨부파일 확인 중…")
            self.pv_files.setVisible(True)
        else:
            self.pv_files.setText("")
            self.pv_files.setVisible(False)
        self.pv_body.setPlainText(message.Preview)
        self._set_buttons(True)

        def run():
            try:
                html = message.HTMLBody
                body = html or message.Body
                attachments = message.attachment_list()
                self.preview_ready.emit(message, body if html else "<pre style='white-space:pre-wrap'>"
                                        + escape(body) + "</pre>", attachments)
            except Exception as e:                       # noqa: BLE001
                logging.error(f"메일 미리보기 실패: {message.Subject} — {e}")
                self.preview_ready.emit(message, f"<p>본문을 받지 못했습니다: {escape(str(e))}</p>", [])
        threading.Thread(target=run, daemon=True).start()

    def _show_preview(self, message, html, attachments):
        if message is not self._current:                # 그 사이 다른 메일을 골랐다
            return
        self._attachments = list(attachments)
        self.pv_body.setHtml(html)
        if self._attachments:
            files_text = " &nbsp;·&nbsp; ".join(
                f"<b>{escape(a.FileName)}</b> ({size_text(a.Size)})" for a in self._attachments
            )
            self.pv_files.setText(f"📎 <b>첨부 ({len(self._attachments)}개):</b> &nbsp;{files_text}")
            self.pv_files.setVisible(True)
        else:
            self.pv_files.setText("")
            self.pv_files.setVisible(False)
        self._set_buttons(True)

    # ── 저장 ──

    def save_eml(self):
        message = self._current
        if message is None:
            return
        path, _ = qtw.QFileDialog.getSaveFileName(self, ".eml 저장", safe_name(message.Subject) + ".eml",
                                                  "메일 (*.eml)")
        if path:
            self._save_in_thread(lambda: (message.SaveAs(path), f"저장했습니다: {path}")[1])

    def save_attachments(self):
        attachments = self._attachments
        if not attachments:
            return
        folder = qtw.QFileDialog.getExistingDirectory(self, "첨부를 저장할 폴더")
        if not folder:
            return

        def work():
            names = []
            for a in attachments:
                path = free_path(Path(folder), safe_name(a.FileName, "attachment"))
                a.SaveAsFile(path)
                names.append(path.name)
            return f"{len(names)}개 저장했습니다: {folder}"
        self._save_in_thread(work)

    def _save_in_thread(self, work):
        self.status.setText("저장 중…")

        def run():
            try:
                self.saved.emit(work())
            except Exception as e:                       # noqa: BLE001
                logging.error(f"메일 저장 실패: {e}")
                self.saved.emit(f"저장하지 못했습니다: {e}")
        threading.Thread(target=run, daemon=True).start()

    def _show_saved(self, text):
        self.status.setText(text)
        logging.info(f"Search 탭 — {text}")
