import asyncio
from enum import Enum
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QPlainTextEdit, QInputDialog,
    QListWidget, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView, QAbstractItemView, QProgressBar,
)

from PySide6.QtCore import Qt, QThread, Slot, Signal
import qdarktheme

import utils


class WorkMode(Enum):
    LoadSpeakers = 1
    Playback = 2
    SaveAudio = 3
    ImportVo = 4


class Worker(QThread):
    sig_load_vo_list = Signal(object)
    sig_send_import_error = Signal(str)
    sig_update_progressbar = Signal(int)

    def __init__(self, mode: WorkMode, speaker=None, text=None, filepath=None, vo_data=None):
        super().__init__()
        self.mode: WorkMode = mode
        self.speaker: str | None = speaker
        self.text: str | None = text
        self.filepath: str | None = filepath
        self.vo_data: list | None = vo_data

    def run(self) -> None:
        if self.mode == WorkMode.LoadSpeakers:
            vo_list = asyncio.run(utils.load_vo_list())
            self.sig_load_vo_list.emit(vo_list)
        if self.mode == WorkMode.Playback:
            asyncio.run(utils.play_back(self.speaker, self.text))
        if self.mode == WorkMode.SaveAudio:
            asyncio.run(utils.save_audio(self.speaker, self.text, self.filepath))
        if self.mode == WorkMode.ImportVo:
            try:
                self.sig_update_progressbar.emit(0)
                # gen vo
                for i, (name, lang, speaker, text, filepath) in enumerate(self.vo_data):
                    asyncio.run(utils.synthesis(speaker, text, str(filepath)))
                    self.sig_update_progressbar.emit(int(i / len(self.vo_data) * 100))
                # import vo
                asyncio.set_event_loop(asyncio.new_event_loop())
                utils.waapi_import_vo(self.vo_data)
            except Exception as e:
                self.sig_send_import_error.emit(str(e))


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.create_widgets()
        self.create_layouts()
        self.create_connections()
        self.speakers = []
        self.workers = []
        self.load_speakers()
        self.setWindowTitle('WAAPI Edge TTS')
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setAcceptDrops(True)
        self.resize(800, 600)

    def create_widgets(self):
        self.edit_search = QLineEdit()
        self.edit_search.setPlaceholderText('Search speaker')
        self.edit_search.setMaximumWidth(250)

        self.speaker_list_w = QListWidget()
        self.speaker_list_w.setMaximumWidth(250)

        self.edit_input = QPlainTextEdit()
        self.edit_input.setPlaceholderText('Input playback text')
        self.edit_input.setMaximumWidth(250)

        self.btn_play_back = QPushButton('Playback')
        self.btn_play_back.setMaximumWidth(250)
        self.btn_save_audio = QPushButton('Save audio')
        self.btn_save_audio.setMaximumWidth(250)

        self.btn_set_speaker = QPushButton('Set speaker')
        self.btn_set_language = QPushButton('Set language')
        self.btn_import_vo = QPushButton('Import voice')

        self.table = QTableWidget()
        self.table.setAcceptDrops(True)
        # self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setRowCount(10)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['VoiceName', 'Language', 'Speaker', 'Text'])

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setVisible(False)

    def create_layouts(self):
        layout = QHBoxLayout()
        _layout = QVBoxLayout()
        _layout.addWidget(self.edit_search)
        _layout.addWidget(self.speaker_list_w)
        _layout.addWidget(self.edit_input)
        _layout.addWidget(self.btn_play_back)
        _layout.addWidget(self.btn_save_audio)
        layout.addLayout(_layout)

        _layout = QVBoxLayout()
        btn_layout = QHBoxLayout()
        btn_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        btn_layout.addWidget(self.btn_set_language)
        btn_layout.addWidget(self.btn_set_speaker)
        btn_layout.addWidget(self.btn_import_vo)
        _layout.addLayout(btn_layout)
        _layout.addWidget(self.table)
        _layout.addWidget(self.progress_bar)
        layout.addLayout(_layout)

        self.setLayout(layout)

    def create_connections(self):
        self.table.dragEnterEvent = self.table_drag_enter_event
        self.table.dragMoveEvent = self.table_drag_move_event
        self.table.dropEvent = self.table_drop_event

        self.edit_search.textChanged.connect(self.search_text_changed)
        self.btn_play_back.clicked.connect(self.play_back)
        self.btn_save_audio.clicked.connect(self.save_audio)

        self.btn_set_speaker.clicked.connect(self.set_speaker)
        self.btn_set_language.clicked.connect(self.set_language)
        self.btn_import_vo.clicked.connect(self.import_vo)

    def table_drag_enter_event(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def table_drag_move_event(self, event):
        if event.mimeData().hasUrls:
            file = event.mimeData().urls()[0].toLocalFile()
            print(file)
            if file.endswith('.xlsx'):
                event.setDropAction(Qt.CopyAction)
                event.accept()
            else:
                event.ignore()
        else:
            event.ignore()

    def table_drop_event(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()
            file = event.mimeData().urls()[0].toLocalFile()
            self.import_xl(file)
        else:
            event.ignore()

    def import_xl(self, file: str):
        data = utils.load_xl(file)
        if not data: return

        self.table.setRowCount(len(data))

        for row, i in enumerate(data):
            name = i.get('VoiceName', None)
            lang = i.get('Language', None)
            speaker = i.get('Speaker', None)
            text = i.get('Text', None)

            if name:
                self.table.setItem(row, 0, QTableWidgetItem(name))
            if lang:
                self.table.setItem(row, 1, QTableWidgetItem(lang))
            if speaker:
                self.table.setItem(row, 2, QTableWidgetItem(speaker))
            if text:
                self.table.setItem(row, 3, QTableWidgetItem(text))

    def load_speakers(self):
        worker = Worker(WorkMode.LoadSpeakers)
        worker.sig_load_vo_list.connect(self.update_speakers)
        worker.start()
        self.workers.append(worker)

    @Slot()
    def update_speakers(self, data):
        self.speakers = data
        self.speakers.sort(key=lambda x: x['ShortName'])

        self.speaker_list_w.clear()
        items = [i['ShortName'] + ', ' + i['Gender'] for i in self.speakers]
        self.speaker_list_w.addItems(items)
        self.speaker_list_w.setCurrentRow(0)

    @Slot()
    def search_text_changed(self, text: str):
        text = text.strip()
        if not self.speakers: return

        self.speaker_list_w.clear()
        items = []
        for i in self.speakers:
            item: str = i['ShortName'] + ', ' + i['Gender']
            if text.lower() in item.lower():
                items.append(item)
        if not items: return

        self.speaker_list_w.addItems(items)
        self.speaker_list_w.setCurrentRow(0)

    @Slot()
    def play_back(self):
        input_text = self.edit_input.toPlainText().strip()
        if not input_text:
            return
        input_text = input_text.replace('\n', '')

        current_speaker = self.speaker_list_w.currentItem().text()
        current_speaker = current_speaker.split(',')[0]

        self.btn_play_back.setDisabled(True)
        worker = Worker(WorkMode.Playback, current_speaker, input_text)
        worker.finished.connect(lambda: self.btn_play_back.setDisabled(False))
        worker.start()
        self.workers.append(worker)

    @Slot()
    def save_audio(self):
        filepath, _ = QFileDialog.getSaveFileName(self, filter='*.wav;;*.mp3')
        if not filepath:
            return

        input_text = self.edit_input.toPlainText().strip()
        if not input_text:
            return
        input_text = input_text.replace('\n', '')

        current_speaker = self.speaker_list_w.currentItem().text()
        current_speaker = current_speaker.split(',')[0]

        self.btn_save_audio.setDisabled(True)
        worker = Worker(
            WorkMode.SaveAudio,
            speaker=current_speaker,
            text=input_text,
            filepath=filepath
        )
        worker.finished.connect(lambda: self.btn_save_audio.setDisabled(False))
        worker.start()
        self.workers.append(worker)

    @Slot()
    def set_speaker(self):
        select = self.table.selectionModel()
        if not select.hasSelection(): return

        current_speaker = self.speaker_list_w.currentItem().text()
        current_speaker = current_speaker.split(',')[0]

        rows = select.selectedIndexes()
        for i in rows:
            if not i.column() == 2: continue
            item = self.table.itemFromIndex(i)
            if not item:
                item = QTableWidgetItem(current_speaker)
                self.table.setItem(i.row(), i.column(), item)
            else:
                item.setText(current_speaker)

    @Slot()
    def set_language(self):
        select = self.table.selectionModel()
        if not select.hasSelection(): return

        lang = self.get_lang_dlg()
        if not lang: return

        rows = select.selectedIndexes()
        for i in rows:
            if not i.column() == 1: continue
            item = self.table.itemFromIndex(i)
            if not item:
                item = QTableWidgetItem(lang)
                self.table.setItem(i.row(), i.column(), item)
            else:
                item.setText(lang)

    def get_lang_dlg(self):
        try:
            items = utils.waapi_get_lang_list()
        except Exception as e:
            QMessageBox.warning(self, 'Error', str(e))
            return
        lang, is_selected = QInputDialog.getItem(self, 'Set language', 'Language:', items, editable=False)
        if is_selected:
            return lang

    @Slot()
    def import_vo(self):
        vo_data = []
        for i in range(self.table.rowCount()):
            name_item = self.table.item(i, 0)
            lang_item = self.table.item(i, 1)
            speaker_item = self.table.item(i, 2)
            text_item = self.table.item(i, 3)
            if not (name_item and lang_item and speaker_item and text_item):
                continue

            name = name_item.text().strip()
            lang = lang_item.text().strip()
            speaker = speaker_item.text().strip()
            text = text_item.text().strip()
            filepath = Path(f'{name}_{lang}.wav')
            filepath = filepath.resolve()
            if not (name and lang and speaker and text):
                continue
            vo_data.append([name, lang, speaker, text, filepath])

        if not vo_data:
            QMessageBox.warning(self, 'Warning', 'Empty vo data')
            return

        self.btn_import_vo.setDisabled(True)
        self.progress_bar.setVisible(True)
        worker = Worker(WorkMode.ImportVo, vo_data=vo_data)
        worker.finished.connect(lambda: (self.btn_import_vo.setDisabled(False),
                                         self.progress_bar.setVisible(False)))
        worker.sig_send_import_error.connect(lambda err: QMessageBox.warning(self, 'Import Vo Error', err))
        worker.sig_update_progressbar.connect(lambda value: self.progress_bar.setValue(value))
        worker.start()
        self.workers.append(worker)

    def clean_local_wav(self):
        for i in Path().glob('*.wav'):
            i.unlink()

    def closeEvent(self, event):
        for i in self.workers:  # type:QThread
            i.terminate()
        # 删除本地生成 wav
        self.clean_local_wav()


app = QApplication()
qdarktheme.setup_theme('auto')
window = Window()
window.show()
app.exec()
