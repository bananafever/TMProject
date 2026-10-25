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


def size_text(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / 1024 / 1024:.1f}MB"
    return f"{max(1, round(size / 1024))}KB"


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
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        row = qtw.QHBoxLayout()
        self.query = qtw.QLineEdit()
        self.query.setPlaceholderText("검색어 — from:이름  subject:제목  body:본문  도 됩니다")
        self.query.setClearButtonEnabled(True)
        self.query.returnPressed.connect(self.search)
        self.folder_combo = qtw.QComboBox()
        self.folder_combo.setMinimumWidth(220)
        self.folder_combo.addItem("전체 메일함", None)
        self.folder_combo.currentIndexChanged.connect(self._folder_changed)
        self.refresh_btn = qtw.QPushButton("↻")
        self.refresh_btn.setObjectName("secondaryButton")
        self.refresh_btn.setToolTip("폴더 목록 새로 받기")
        self.refresh_btn.clicked.connect(lambda: self.load_folders(refresh=True))
        self.search_btn = qtw.QPushButton("검색")
        self.search_btn.clicked.connect(self.search)
        row.addWidget(self.query, 1)
        row.addWidget(self.folder_combo)
        row.addWidget(self.refresh_btn)
        row.addWidget(self.search_btn)
        layout.addLayout(row)

        row = qtw.QHBoxLayout()
        self.subfolders = qtw.QCheckBox("하위 폴더 포함")
        self.subfolders.setChecked(True)
        self.subfolders.setEnabled(False)
        self.period = qtw.QComboBox()
        for label, delta in PERIODS:
            self.period.addItem(label, delta)
        self.period.setCurrentIndex(3)
        self.has_attachments = qtw.QCheckBox("첨부 있음")
        self.sender = qtw.QLineEdit()
        self.sender.setPlaceholderText("보낸 사람")
        self.sender.setClearButtonEnabled(True)
        self.sender.setFixedWidth(160)
        self.sender.returnPressed.connect(self.search)
        self.status = qtw.QLabel("")
        self.status.setObjectName("settingsHint")
        for w in (self.subfolders, self.period, self.has_attachments, self.sender):
            row.addWidget(w)
        row.addStretch()
        row.addWidget(self.status)
        layout.addLayout(row)

        self.table = qtw.QTableWidget(0, len(COLUMNS))
        self.table.setObjectName("mailTable")
        self.table.setHorizontalHeaderLabels(COLUMNS)
        header = self.table.horizontalHeader()
        font = qtw.QApplication.font()                   # main.use_app_font 와 같은 까닭 — 표·헤더 힌팅
        self.table.setFont(font)
        header.setFont(font)
        header.setSectionResizeMode(2, qtw.QHeaderView.Stretch)
        for col, width in ((0, 128), (1, 120), (3, 140), (4, 44)):
            header.setSectionResizeMode(col, qtw.QHeaderView.Interactive)
            header.resizeSection(col, width)
        header.setHighlightSections(False)
        header.setSortIndicator(0, qtc.Qt.DescendingOrder)          # 받은 날 새 것부터
        header.setDefaultAlignment(qtc.Qt.AlignLeft | qtc.Qt.AlignVCenter)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(qtw.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(qtw.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(qtw.QAbstractItemView.NoEditTriggers)
        self.table.currentCellChanged.connect(lambda row, *_: self._select(row))

        self.preview = qtw.QFrame()
        self.preview.setObjectName("settingsSection")
        box = qtw.QVBoxLayout(self.preview)
        box.setContentsMargins(14, 10, 14, 10)
        box.setSpacing(6)
        top = qtw.QHBoxLayout()
        self.pv_subject = qtw.QLabel("메일을 고르면 여기에 보입니다")
        self.pv_subject.setObjectName("sectionTitle")
        self.pv_subject.setTextInteractionFlags(qtc.Qt.TextSelectableByMouse)
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
        self.pv_meta = qtw.QLabel("")
        self.pv_meta.setObjectName("settingsHint")
        self.pv_meta.setWordWrap(True)
        self.pv_meta.setTextInteractionFlags(qtc.Qt.TextSelectableByMouse)
        self.pv_files = qtw.QLabel("")
        self.pv_files.setWordWrap(True)
        self.pv_body = qtw.QTextBrowser()
        self.pv_body.setObjectName("previewBody")
        self.pv_body.setOpenExternalLinks(True)
        self.pv_body.setOpenLinks(True)
        box.addWidget(self.pv_meta)
        box.addWidget(self.pv_files)
        box.addWidget(self.pv_body, 1)
        self._set_buttons(False)

        splitter = qtw.QSplitter(qtc.Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self.table)
        splitter.addWidget(self.preview)
        splitter.setSizes([260, 300])
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
                self.folders_ready.emit(self.get_mailbox().folders(refresh=refresh))
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
            cells = [received.strftime("%Y-%m-%d %H:%M") if received else "", message.SenderName, message.Subject,
                     path.rsplit(" / ", 1)[-1], "●" if message.data.get("hasAttachments") else ""]
            for c, text in enumerate(cells):
                item = qtw.QTableWidgetItem(text)
                if c == 0:
                    item.setData(qtc.Qt.UserRole, message)
                elif c == 2:
                    item.setToolTip(message.Preview or message.Subject)
                elif c == 3:
                    item.setToolTip(path)
                elif c == 4:
                    item.setTextAlignment(qtc.Qt.AlignCenter)
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

    def _clear_preview(self, title="메일을 고르면 여기에 보입니다"):
        self._current, self._attachments = None, None
        self.pv_subject.setText(title)
        self.pv_meta.setText("")
        self.pv_files.setText("")
        self.pv_body.clear()
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
        meta = [x for x in (message.SenderName, received, path) if x]
        if message.To:
            meta.append("받는 사람 " + message.To)
        self.pv_subject.setText(message.Subject or "(제목 없음)")
        self.pv_meta.setText(" · ".join(meta))
        self.pv_files.setText("첨부 확인 중…" if message.data.get("hasAttachments") else "")
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
        self.pv_files.setText(("첨부  " + "  ·  ".join(f"{a.FileName} ({size_text(a.Size)})" for a in self._attachments))
                             if self._attachments else "")
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
