import sys
import os
import json
import time
import webbrowser
from datetime import datetime

# PyQt5
from PyQt5.QtCore import Qt, QUrl, QTimer
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog,
    QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem, QGridLayout, QTextEdit, QSpinBox, QSlider
)
from PyQt5.QtMultimediaWidgets import QVideoWidget

import keyboard

try:
    from obswebsocket import obsws, requests, events
except ImportError:
    print("obs-websocket-py がインストールされていません。")
    sys.exit(1)

import vlc  # python-vlc のインポート

# ヘルパー関数: 秒数を時間・分・秒に変換
def format_duration(seconds):
    seconds = int(round(seconds))
    hrs = seconds // 3600
    mins = (seconds % 3600) // 60
    secs = seconds % 60
    parts = []
    if hrs > 0:
        parts.append(f"{hrs}時間")
    if mins > 0 or hrs > 0:
        parts.append(f"{mins}分")
    parts.append(f"{secs}秒")
    return "".join(parts)

# ベースディレクトリの取得
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_pixmap(filename):
    path = os.path.join(BASE_DIR, filename)
    return QPixmap(path)
    
# ConfigManager

class ConfigManager:
    CONFIG_FILE = "config.json"

    def __init__(self):
        self.data = {
            "obs_host": "localhost",
            "obs_port": 4455,
            "obs_password": "",
            "projects_path": "projects",
            "hotkey": "V",
            "last_connection_success": False
        }
        self.local_config_path = os.path.abspath(self.CONFIG_FILE)
        self.fallback_config_dir = os.path.join(os.path.expanduser("~"), ".obs_timestamp_app_globalhotkey")
        self.fallback_config_path = os.path.join(self.fallback_config_dir, "config.json")
        self.load_config()

    def load_config(self):
        loaded_path = None
        if os.path.exists(self.local_config_path):
            try:
                with open(self.local_config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                self.data.update(loaded)
                loaded_path = self.local_config_path
            except Exception as e:
                print(f"Failed to load config from {self.local_config_path}:", e)
        elif os.path.exists(self.fallback_config_path):
            try:
                with open(self.fallback_config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                self.data.update(loaded)
                loaded_path = self.fallback_config_path
            except Exception as e:
                print(f"Failed to load config from {self.fallback_config_path}:", e)

        if loaded_path:
            print(f"Loaded config from {loaded_path}")
        else:
            print("No config file found, using default settings.")

    def save_config(self):
        try:
            with open(self.local_config_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
            print(f"Saved config to {self.local_config_path}")
            return
        except PermissionError:
            print(f"PermissionError: cannot write {self.local_config_path}, fallback to home directory.")
        except Exception as e:
            print(f"Failed to save config.json to local path: {e}, fallback to home directory.")

        try:
            os.makedirs(self.fallback_config_dir, exist_ok=True)
            with open(self.fallback_config_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
            print(f"Saved config to {self.fallback_config_path}")
        except Exception as e:
            print(f"Failed to save config.json to fallback path: {e}")


# Project

class Project:
    def __init__(self, name, video_file_path=None, timestamps=None):
        self.name = name
        self.video_file_path = video_file_path if video_file_path else ""
        self.timestamps = timestamps if timestamps else []

    def add_timestamp(self, sec):
        self.timestamps.append({"sec": sec, "note": ""})

    def set_note(self, index, note):
        if 0 <= index < len(self.timestamps):
            self.timestamps[index]["note"] = note

    def to_dict(self):
        return {
            "name": self.name,
            "video_file_path": self.video_file_path,
            "timestamps": self.timestamps
        }

    @staticmethod
    def from_dict(dct):
        return Project(
            name=dct["name"],
            video_file_path=dct.get("video_file_path", ""),
            timestamps=dct.get("timestamps", [])
        )

    def save_json(self, folder_path):
        os.makedirs(folder_path, exist_ok=True)
        file_path = os.path.join(folder_path, f"{self.name}.json")
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=4)
            print(f"[DEBUG] Saved project JSON: {file_path}, video_file_path={self.video_file_path}")
        except Exception as e:
            print("[ERROR] Failed to save project JSON:", e)

    @staticmethod
    def load_json(folder_path, name):
        file_path = os.path.join(folder_path, f"{name}.json")
        if not os.path.exists(file_path):
            return None
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                dct = json.load(f)
            return Project.from_dict(dct)
        except Exception as e:
            print(f"[ERROR] Failed to load project JSON '{file_path}':", e)
            return None

    @staticmethod
    def list_projects(folder_path):
        if not os.path.exists(folder_path):
            return []
        files = os.listdir(folder_path)
        projects = []
        for f in files:
            if f == "config.json":
                continue
            if f.endswith(".json"):
                projects.append(f[:-5])
        return projects


# OBSController

class OBSController:
    def __init__(self, host, port, password, main_app):
        self.host = host
        self.port = port
        self.password = password
        self.main_app = main_app
        self.ws = None
        self.is_connected = False
        self.last_known_stop_path = None

    def connect(self):
        try:
            self.ws = obsws(self.host, self.port, self.password)
            self.ws.connect()
            self.is_connected = True
            self.ws.register(self.on_record_state_changed_5x, events.RecordStateChanged)
            self.ws.register(self.on_recording_started_4x, events.RecordingStarted)
            self.ws.register(self.on_recording_stopped_4x, events.RecordingStopped)
            return True
        except Exception as e:
            print("OBS connection failed:", e)
            self.is_connected = False
            return False

    def disconnect(self):
        if self.ws and self.is_connected:
            self.ws.disconnect()
            self.is_connected = False

    def on_record_state_changed_5x(self, event):
        data = event.datain
        if not data:
            return
        output_state = data.get("outputState", "")
        output_path = data.get("outputPath", "")
        print(f"[DEBUG][5.x event] state={output_state}, path={output_path}")
        if output_state == "OBS_WEBSOCKET_OUTPUT_STOPPED":
            self.last_known_stop_path = output_path

    def on_recording_started_4x(self, event):
        filename = ""
        try:
            filename = event.getRecordingFilename()
        except:
            pass
        print(f"[DEBUG][4.x event] RecordingStarted, file={filename}")

    def on_recording_stopped_4x(self, event):
        filename = ""
        try:
            filename = event.getRecordingFilename()
        except:
            pass
        print(f"[DEBUG][4.x event] RecordingStopped, file={filename}")
        self.last_known_stop_path = filename

    def get_final_record_path(self):
        if self.last_known_stop_path:
            print("[DEBUG] Using last_known_stop_path from event:", self.last_known_stop_path)
            return self.last_known_stop_path
        if not self.is_connected or not self.ws:
            return ""
        path = ""
        try:
            status = self.ws.call(requests.GetRecordStatus())
            if hasattr(status, "datain"):
                data = status.datain
                path = data.get("outputPath", "")
                print("[DEBUG] final path (GetRecordStatus 5.x):", path)
        except:
            pass
        if not path:
            try:
                old_status = self.ws.call(requests.GetRecordingStatus())
                path = old_status.getRecordingFilename()
                print("[DEBUG] final path (GetRecordingStatus 4.x):", path)
            except:
                print("[ERROR] Could not retrieve final path via 4.x either.")
        return path

    def clear_stop_path(self):
        self.last_known_stop_path = None


# SetupDialog

class SetupDialog(QDialog):
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("初期設定 - OBS 接続情報")
        self.config_manager = config_manager

        logo_label = QLabel()
        pixmap = load_pixmap("logo.png")
        if not pixmap.isNull():
            # ロゴ画像を高さ300に設定
            logo_label.setPixmap(pixmap.scaledToHeight(300, Qt.SmoothTransformation))

        layout = QVBoxLayout()
        layout.addWidget(logo_label)

        self.host_edit = QLineEdit()
        self.host_edit.setText(self.config_manager.data["obs_host"])
        self.port_edit = QLineEdit()
        self.port_edit.setText(str(self.config_manager.data["obs_port"]))
        self.password_edit = QLineEdit()
        self.password_edit.setText(self.config_manager.data["obs_password"])
        self.password_edit.setEchoMode(QLineEdit.Password)

        self.project_path_edit = QLineEdit()
        self.project_path_edit.setText(self.config_manager.data["projects_path"])
        browse_button = QPushButton("参照")
        browse_button.clicked.connect(self.browse_project_path)

        form_layout = QGridLayout()
        form_layout.addWidget(QLabel("OBS ホスト:"), 0, 0)
        form_layout.addWidget(self.host_edit, 0, 1)
        form_layout.addWidget(QLabel("OBS ポート:"), 1, 0)
        form_layout.addWidget(self.port_edit, 1, 1)
        form_layout.addWidget(QLabel("OBS パスワード:"), 2, 0)
        form_layout.addWidget(self.password_edit, 2, 1)
        form_layout.addWidget(QLabel("プロジェクト保存フォルダ:"), 3, 0)
        form_layout.addWidget(self.project_path_edit, 3, 1)
        form_layout.addWidget(browse_button, 3, 2)
        layout.addLayout(form_layout)

        self.connect_button = QPushButton("接続テスト")
        self.connect_button.clicked.connect(self.test_connection)
        layout.addWidget(self.connect_button)

        self.setLayout(layout)
        self.resize(400, 200)

    def browse_project_path(self):
        dir_path = QFileDialog.getExistingDirectory(self, "プロジェクト保存フォルダを選択")
        if dir_path:
            self.project_path_edit.setText(dir_path)

    def test_connection(self):
        host = self.host_edit.text().strip()
        port_text = self.port_edit.text().strip()
        password = self.password_edit.text().strip()

        if not port_text.isdigit():
            QMessageBox.critical(self, "エラー", "ポートには数値を入力してください。")
            return
        port = int(port_text)

        self.config_manager.data["obs_host"] = host
        self.config_manager.data["obs_port"] = port
        self.config_manager.data["obs_password"] = password
        self.config_manager.data["projects_path"] = self.project_path_edit.text()
        self.config_manager.save_config()

        test_ws = obsws(host, port, password)
        try:
            test_ws.connect()
            test_ws.disconnect()
            QMessageBox.information(self, "成功", "OBS に正常に接続できました。")
            self.config_manager.data["last_connection_success"] = True
            self.config_manager.save_config()
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"OBS への接続に失敗しました。設定を見直してください。\n{e}")


# MainMenuWindow

class MainMenuWindow(QMainWindow):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.config_manager = main_app.config_manager

        # ウィンドウタイトルに「TimePortal」を追加
        self.setWindowTitle("TimePortal - メインメニュー")
        self.resize(600, 400)

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        self.logo_label = QLabel()
        logo_pixmap = load_pixmap("logo.png")
        if not logo_pixmap.isNull():
            # ロゴ画像を高さ300に設定
            self.logo_label.setPixmap(logo_pixmap.scaledToHeight(300, Qt.SmoothTransformation))
        top_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)
        layout.addLayout(top_layout)

        layout.addWidget(QLabel("プロジェクト一覧:"))
        self.project_list_widget = QListWidget()
        layout.addWidget(self.project_list_widget)
        self.refresh_project_list()

        open_button = QPushButton("選択したプロジェクトを編集画面で開く")
        open_button.clicked.connect(self.open_selected_project)
        layout.addWidget(open_button)

        create_button = QPushButton("新規プロジェクトを作成")
        create_button.clicked.connect(self.create_new_project)
        layout.addWidget(create_button)

        hotkey_layout = QHBoxLayout()
        hotkey_layout.addWidget(QLabel("タイムスタンプ用ホットキー:"))
        self.hotkey_edit = QLineEdit()
        self.hotkey_edit.setText(self.config_manager.data["hotkey"])
        hotkey_layout.addWidget(self.hotkey_edit)

        hotkey_set_button = QPushButton("設定")
        hotkey_set_button.clicked.connect(self.set_hotkey)
        hotkey_layout.addWidget(hotkey_set_button)
        layout.addLayout(hotkey_layout)

        icon_layout = QHBoxLayout()
        icon_layout.addStretch(1)
        self.icon_label = QLabel()
        icon_pixmap = load_pixmap("icon.png")
        if not icon_pixmap.isNull():
            # アイコン画像は従来サイズ（高さ120のまま）
            self.icon_label.setPixmap(icon_pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.icon_label.mousePressEvent = self.open_dev_url
        icon_layout.addWidget(self.icon_label, alignment=Qt.AlignRight)
        layout.addLayout(icon_layout)

        self.setCentralWidget(central_widget)

    def open_dev_url(self, event):
        webbrowser.open("https://lit.link/shake1227")

    def refresh_project_list(self):
        self.project_list_widget.clear()
        project_names = Project.list_projects(self.config_manager.data["projects_path"])
        for name in project_names:
            self.project_list_widget.addItem(name)

    def open_selected_project(self):
        item = self.project_list_widget.currentItem()
        if not item:
            return
        project_name = item.text()
        self.main_app.open_project_edit(project_name)

    def create_new_project(self):
        text, ok = QInputDialogWithTitle.getText(self, "新規プロジェクト名", "プロジェクト名を入力:")
        if ok and text.strip():
            self.main_app.create_and_open_project(text.strip())
            self.refresh_project_list()

    def set_hotkey(self):
        new_hotkey = self.hotkey_edit.text().strip()
        if not new_hotkey:
            return
        if self.main_app.current_hotkey_handler is not None:
            keyboard.remove_hotkey(self.main_app.current_hotkey_handler)
        self.main_app.current_hotkey_handler = keyboard.add_hotkey(new_hotkey, self.main_app.global_hotkey_callback)
        self.config_manager.data["hotkey"] = new_hotkey
        self.config_manager.save_config()
        QMessageBox.information(self, "ホットキー", f"ホットキーを '{new_hotkey}' に変更しました。")


# RecordingWindow

class RecordingWindow(QMainWindow):
    def __init__(self, main_app, project: Project):
        super().__init__()
        self.main_app = main_app
        self.project = project
        self.setWindowTitle(f"録画中 - プロジェクト: {project.name}")
        self.resize(600, 400)

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)

        self.logo_label = QLabel()
        pixmap = load_pixmap("logo.png")
        if not pixmap.isNull():
            # ロゴ画像を高さ300に設定
            self.logo_label.setPixmap(pixmap.scaledToHeight(300, Qt.SmoothTransformation))
        layout.addWidget(self.logo_label, alignment=Qt.AlignTop)

        self.timestamp_list_widget = QListWidget()
        layout.addWidget(QLabel("タイムスタンプ一覧:"))
        layout.addWidget(self.timestamp_list_widget)

        icon_layout = QHBoxLayout()
        icon_layout.addStretch(1)
        self.icon_label = QLabel()
        icon_pixmap = load_pixmap("icon.png")
        if not icon_pixmap.isNull():
            # アイコン画像は従来サイズ（高さ120のまま）
            self.icon_label.setPixmap(icon_pixmap.scaledToHeight(120, Qt.SmoothTransformation))
        self.icon_label.mousePressEvent = self.open_dev_url
        icon_layout.addWidget(self.icon_label, alignment=Qt.AlignRight)
        layout.addLayout(icon_layout)

        self.setCentralWidget(central_widget)

        self.recording_start_time = time.time()
        self.update_timestamp_list()

    def open_dev_url(self, event):
        webbrowser.open("https://lit.link/shake1227")

    def add_timestamp(self):
        current_sec = time.time() - self.recording_start_time
        self.project.add_timestamp(round(current_sec, 2))
        self.update_timestamp_list()

    def update_timestamp_list(self):
        self.timestamp_list_widget.clear()
        for ts in self.project.timestamps:
            sec = ts["sec"]
            formatted_time = format_duration(sec)
            self.timestamp_list_widget.addItem(f"{formatted_time}")

    def closeEvent(self, event):
        self.project.save_json(self.main_app.config_manager.data["projects_path"])
        super().closeEvent(event)


# EditWindow

class EditWindow(QMainWindow):
    def __init__(self, main_app, project: Project):
        super().__init__()
        self.main_app = main_app
        self.project = project
        self.setWindowTitle(f"編集画面 - プロジェクト: {project.name}")
        self.resize(1000, 600)

        central_widget = QWidget()
        main_layout = QHBoxLayout(central_widget)

        # 左側レイアウト（動画とコントロール）
        left_layout = QVBoxLayout()

        self.logo_label = QLabel()
        pixmap = load_pixmap("logo.png")
        if not pixmap.isNull():
            # ロゴ画像を高さ300に設定
            self.logo_label.setPixmap(pixmap.scaledToHeight(300, Qt.SmoothTransformation))
        left_layout.addWidget(self.logo_label, alignment=Qt.AlignTop)

        self.video_widget = QVideoWidget()
        self.vlc_instance = vlc.Instance('--avcodec-hw=none', '--no-video-title-show')
        self.vlc_player = self.vlc_instance.media_player_new()
        self.vlc_player.set_hwnd(int(self.video_widget.winId()))
        self.video_widget.setMinimumSize(800, 450)  # 16:9の最小サイズ
        left_layout.addWidget(self.video_widget)

        self.position_slider = QSlider(Qt.Horizontal)
        left_layout.addWidget(self.position_slider)
        self.position_slider.sliderMoved.connect(self.on_slider_moved)

        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_slider)
        self.update_timer.start(500)

        self.time_label = QLabel("00:00 / 00:00")
        left_layout.addWidget(self.time_label)

        control_layout = QHBoxLayout()
        play_btn = QPushButton("▶ 再生")
        pause_btn = QPushButton("■ 停止")
        backward_10_btn = QPushButton("<< 10秒戻し")
        forward_10_btn = QPushButton("10秒送り >>")
        speed_label = QLabel("再生速度:")
        self.speed_spinbox = QSpinBox()
        self.speed_spinbox.setRange(1, 5)
        self.speed_spinbox.setValue(1)
        control_layout.addWidget(play_btn)
        control_layout.addWidget(pause_btn)
        control_layout.addWidget(backward_10_btn)
        control_layout.addWidget(forward_10_btn)
        control_layout.addWidget(speed_label)
        control_layout.addWidget(self.speed_spinbox)
        left_layout.addLayout(control_layout)

        main_layout.addLayout(left_layout, stretch=2)

        right_layout = QVBoxLayout()
        self.timestamp_list_widget = QListWidget()
        for ts in self.project.timestamps:
            sec = ts["sec"]
            note = ts.get("note", "")
            formatted_time = format_duration(sec)
            self.timestamp_list_widget.addItem(f"{formatted_time}: {note}")
        self.timestamp_list_widget.itemDoubleClicked.connect(self.on_timestamp_double_clicked)
        right_layout.addWidget(QLabel("タイムスタンプ一覧:"))
        right_layout.addWidget(self.timestamp_list_widget)

        self.note_edit = QTextEdit()
        right_layout.addWidget(QLabel("選択したタイムスタンプのメモ:"))
        right_layout.addWidget(self.note_edit)

        copy_button = QPushButton("全タイムスタンプ+メモをコピー")
        right_layout.addWidget(copy_button)

        icon_layout = QHBoxLayout()
        icon_layout.addStretch(1)
        self.icon_label = QLabel()
        icon_pixmap = load_pixmap("icon.png")
        if not icon_pixmap.isNull():
            # アイコン画像は従来サイズ（高さ120のまま）
            self.icon_label.setPixmap(icon_pixmap.scaledToHeight(120, Qt.SmoothTransformation))
        self.icon_label.mousePressEvent = self.open_dev_url
        icon_layout.addWidget(self.icon_label, alignment=Qt.AlignRight)
        right_layout.addLayout(icon_layout)

        main_layout.addLayout(right_layout, stretch=1)
        self.setCentralWidget(central_widget)

        self.timestamp_list_widget.currentRowChanged.connect(self.on_timestamp_selected)
        self.note_edit.textChanged.connect(self.on_note_changed)
        play_btn.clicked.connect(self.on_play)
        pause_btn.clicked.connect(self.on_pause)
        backward_10_btn.clicked.connect(self.on_backward_10)
        forward_10_btn.clicked.connect(self.on_forward_10)
        self.speed_spinbox.valueChanged.connect(self.on_speed_changed)
        copy_button.clicked.connect(self.copy_timestamps_and_notes)

        self.load_video()

    def load_video(self):
        video_path = os.path.abspath(self.project.video_file_path)
        if os.path.exists(video_path):
            media = self.vlc_instance.media_new(video_path)
            self.vlc_player.set_media(media)
            self.vlc_player.play()
            QTimer.singleShot(200, self.vlc_player.pause)
        else:
            print(f"[WARN] Video file not found: {video_path}")

    def update_slider(self):
        length = self.vlc_player.get_length()
        if length > 0:
            self.position_slider.setMaximum(length)
        current_time = self.vlc_player.get_time()
        if current_time >= 0:
            self.position_slider.blockSignals(True)
            self.position_slider.setValue(current_time)
            self.position_slider.blockSignals(False)
        total_sec = length // 1000 if length > 0 else 0
        current_sec = current_time // 1000 if current_time >= 0 else 0
        self.time_label.setText(f"{current_sec//60:02d}:{current_sec%60:02d} / {total_sec//60:02d}:{total_sec%60:02d}")

    def on_slider_moved(self, position):
        self.vlc_player.set_time(position)

    def open_dev_url(self, event):
        webbrowser.open("https://lit.link/shake1227")

    def on_timestamp_selected(self, row):
        if row < 0 or row >= len(self.project.timestamps):
            return
        ts = self.project.timestamps[row]
        note = ts.get("note", "")
        self.note_edit.blockSignals(True)
        self.note_edit.setText(note)
        self.note_edit.blockSignals(False)
        target_ms = int(ts["sec"] * 1000)
        self.vlc_player.set_time(target_ms)
        self.vlc_player.play()

    def on_timestamp_double_clicked(self, item):
        row = self.timestamp_list_widget.row(item)
        if row < 0 or row >= len(self.project.timestamps):
            return
        ts = self.project.timestamps[row]
        target_ms = int(ts["sec"] * 1000)
        self.vlc_player.set_time(target_ms)
        self.vlc_player.play()

    def on_note_changed(self):
        row = self.timestamp_list_widget.currentRow()
        if row < 0 or row >= len(self.project.timestamps):
            return
        note = self.note_edit.toPlainText()
        self.project.set_note(row, note)
        sec = self.project.timestamps[row]["sec"]
        formatted_time = format_duration(sec)
        self.timestamp_list_widget.item(row).setText(f"{formatted_time}: {note}")

    def on_play(self):
        print("[DEBUG] Play button clicked")
        self.vlc_player.play()

    def on_pause(self):
        print("[DEBUG] Pause button clicked")
        self.vlc_player.pause()

    def on_backward_10(self):
        print("[DEBUG] Backward 10s button clicked")
        current_time = self.vlc_player.get_time()
        self.vlc_player.set_time(max(0, current_time - 10000))

    def on_forward_10(self):
        print("[DEBUG] Forward 10s button clicked")
        current_time = self.vlc_player.get_time()
        duration = self.vlc_player.get_length()
        self.vlc_player.set_time(min(duration, current_time + 10000))

    def on_speed_changed(self, value):
        print(f"[DEBUG] Speed changed to {value}")
        try:
            self.vlc_player.set_rate(value)
        except Exception as e:
            QMessageBox.warning(self, "注意", f"再生速度変更に失敗: {e}")

    def copy_timestamps_and_notes(self):
        lines = []
        for ts in self.project.timestamps:
            sec = ts["sec"]
            note = ts.get("note", "")
            formatted_time = format_duration(sec)
            lines.append(f"{formatted_time}: {note}")
        text = "\n".join(lines)
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "コピー", "タイムスタンプとメモをクリップボードにコピーしました。")

    def closeEvent(self, event):
        self.project.save_json(self.main_app.config_manager.data["projects_path"])
        super().closeEvent(event)

# QInputDialogWithTitle

class QInputDialogWithTitle:
    @staticmethod
    def getText(parent, title, label):
        dlg = QDialog(parent)
        dlg.setWindowTitle(title)
        layout = QVBoxLayout(dlg)
        label_widget = QLabel(label)
        layout.addWidget(label_widget)
        line_edit = QLineEdit()
        layout.addWidget(line_edit)
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        cancel_btn = QPushButton("キャンセル")
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        ret_data = {"text": "", "accepted": False}
        def on_ok():
            ret_data["text"] = line_edit.text()
            ret_data["accepted"] = True
            dlg.accept()
        def on_cancel():
            dlg.reject()
        ok_btn.clicked.connect(on_ok)
        cancel_btn.clicked.connect(on_cancel)
        dlg.exec_()
        return ret_data["text"], ret_data["accepted"]

# MainApp

class MainApp(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.setApplicationName("TimePortal")
        self.setWindowIcon(QIcon(os.path.join(BASE_DIR, "task.png")))
        
        # 追加: VLCの存在チェック
        try:
            self.vlc_instance = vlc.Instance('--avcodec-hw=none', '--no-video-title-show')
        except Exception as e:
            QMessageBox.critical(None, "VLC エラー", "VLCがインストールされていません。VLCをインストールしてください。ダウンロードページを開きます。")
            webbrowser.open("https://www.videolan.org/vlc/index.ja.html")
            sys.exit(1)

        self.config_manager = ConfigManager()
        if not self.config_manager.data["last_connection_success"]:
            while True:
                setup_dlg = SetupDialog(self.config_manager)
                result = setup_dlg.exec_()
                if result == QDialog.Accepted:
                    break
                else:
                    reply = QMessageBox.question(
                        None,
                        "確認",
                        "設定が完了していないため終了しますか？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        sys.exit(0)
        self.obs_controller = None
        self.connect_to_obs()
        self.main_menu = MainMenuWindow(self)
        self.main_menu.show()
        self.is_recording = False
        self.current_recording_project = None
        self.recording_window = None
        self.current_hotkey_handler = None
        self.setup_global_hotkey()
        self.record_status_timer = QTimer(self)
        self.record_status_timer.timeout.connect(self.check_record_status)
        self.record_status_timer.start(1000)

    def setup_global_hotkey(self):
        hotkey = self.config_manager.data["hotkey"]
        self.current_hotkey_handler = keyboard.add_hotkey(hotkey, self.global_hotkey_callback)

    def global_hotkey_callback(self):
        if self.is_recording and self.current_recording_project and self.recording_window:
            self.recording_window.add_timestamp()

    def connect_to_obs(self):
        if self.obs_controller and self.obs_controller.is_connected:
            self.obs_controller.disconnect()
        host = self.config_manager.data["obs_host"]
        port = self.config_manager.data["obs_port"]
        password = self.config_manager.data["obs_password"]
        self.obs_controller = OBSController(host, port, password, self)
        if not self.obs_controller.connect():
            QMessageBox.critical(None, "OBS接続エラー", "OBS への接続に失敗しました。再度設定を行ってください。")
            self.config_manager.data["last_connection_success"] = False
            self.config_manager.save_config()
            while True:
                setup_dlg = SetupDialog(self.config_manager)
                result = setup_dlg.exec_()
                if result == QDialog.Accepted:
                    if self.obs_controller.connect():
                        self.config_manager.data["last_connection_success"] = True
                        self.config_manager.save_config()
                        break
                else:
                    reply = QMessageBox.question(
                        None, "確認",
                        "設定が完了していないため終了しますか？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        sys.exit(0)

    def check_record_status(self):
        if not self.obs_controller or not self.obs_controller.is_connected:
            return
        currently_recording = False
        rec_filename = ""
        try:
            status = self.obs_controller.ws.call(requests.GetRecordStatus())
            if hasattr(status, "datain"):
                data = status.datain
                currently_recording = data.get("outputActive", False)
                rec_filename = data.get("outputPath", "")
        except Exception as e:
            print("GetRecordStatus failed, fallback to 4.x old API:", e)
            try:
                old_status = self.obs_controller.ws.call(requests.GetRecordingStatus())
                currently_recording = old_status.getIsRecording()
                rec_filename = old_status.getRecordingFilename()
            except Exception as e2:
                print("GetRecordingStatus also failed:", e2)
                return
        if currently_recording and not self.is_recording:
            self.is_recording = True
            fallback_name = f"Recording_{int(time.time())}"
            project = Project(name=fallback_name, video_file_path="")
            project.save_json(self.config_manager.data["projects_path"])
            self.current_recording_project = project
            if self.recording_window:
                self.recording_window.close()
            self.recording_window = RecordingWindow(self, project)
            self.recording_window.show()
        elif not currently_recording and self.is_recording:
            self.is_recording = False
            self._retrieve_final_path_with_retry(5)

    def _retrieve_final_path_with_retry(self, attempts_left):
        if not self.current_recording_project:
            return
        final_path = self.obs_controller.get_final_record_path()
        if final_path:
            self._finalize_with_path(final_path)
        else:
            if attempts_left > 1:
                print(f"[INFO] final path is empty. retry in 1s... (remaining={attempts_left-1})")
                QTimer.singleShot(1000, lambda: self._retrieve_final_path_with_retry(attempts_left - 1))
            else:
                print("[WARN] Could not retrieve final path after multiple attempts. Fallback name only.")
                self._finalize_with_path("")

    def _finalize_with_path(self, final_path):
        if not self.current_recording_project:
            return
        if final_path:
            base_name = os.path.basename(final_path)
            if base_name:
                old_name = self.current_recording_project.name
                self.current_recording_project.name = base_name
                self.current_recording_project.video_file_path = final_path
                old_file = os.path.join(self.config_manager.data["projects_path"], f"{old_name}.json")
                if os.path.exists(old_file):
                    try:
                        os.remove(old_file)
                    except:
                        pass
                self.current_recording_project.save_json(self.config_manager.data["projects_path"])
                print(f"[DEBUG] final path used: {final_path}")
            else:
                self.current_recording_project.video_file_path = final_path
                self.current_recording_project.save_json(self.config_manager.data["projects_path"])
                print("[DEBUG] final path is empty base_name, using fallback only.")
        else:
            self.current_recording_project.save_json(self.config_manager.data["projects_path"])
        if self.recording_window:
            self.recording_window.close()
            self.recording_window = None
        self.open_project_edit(self.current_recording_project.name)
        self.obs_controller.clear_stop_path()
        self.current_recording_project = None

    def create_and_open_project(self, project_name):
        project = Project(project_name)
        project.save_json(self.config_manager.data["projects_path"])
        if self.recording_window:
            self.recording_window.close()
        self.recording_window = RecordingWindow(self, project)
        self.recording_window.show()
        self.current_recording_project = project

    def open_project_edit(self, project_name):
        proj = Project.load_json(self.config_manager.data["projects_path"], project_name)
        if not proj:
            QMessageBox.warning(None, "エラー", f"プロジェクト '{project_name}' が見つかりません。")
            return
        edit_win = EditWindow(self, proj)
        edit_win.show()

def main():
    app = MainApp(sys.argv)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
