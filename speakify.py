import sys
import json
import os
import threading
import time
import xlwings as xw
import qtawesome as qta
import speech_recognition as sr
# --- UPDATED IMPORTS FOR ADVANCED UI ---
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget,
                               QVBoxLayout, QHBoxLayout, QPushButton,
                               QLabel, QFrame, QComboBox, QStackedWidget,
                               QGridLayout, QMessageBox, QTextEdit, QButtonGroup,
                               QLineEdit, QListWidget, QCompleter) # QLineEdit, QListWidget, QCompleter add kiya gaya hai
from PySide6.QtCore import Qt, QSize, Signal, QObject, QTimer, QPoint, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QFontDatabase, QPalette, QColor, QIcon, QPixmap

EXCEL_FILES_FOLDER_PATH = r"C:\Users\HP\OneDrive\Desktop\IMAGE FORMAT"

STYLESHEET = """
    /* --- Base Font --- */
    QMainWindow {
        font-family: "Segoe UI", sans-serif;
        font-size: 16px;
    }

    /* --- Main Window & Sidebar --- */
    #MainWindow { 
        background-color: #0a0a0f;
        color: #ffffff;
    }
    #ContentWindow {
        background-color: #0f0f15;
        border: none;
    }
    #Sidebar { 
        background-color: rgba(15, 15, 25, 0.95);
        border-right: 1px solid #2a2a35;
    }
    #SidebarTitle {
        font-family: "Audiowide";
        font-size: 24px;
        color: #ffffff;
        margin: 10px 0;
        text-shadow: 0 0 10px rgba(0, 123, 255, 0.7);
    }
    #SidebarButton {
        font-size: 18px;
        font-weight: 500;
        text-align: left;
        padding: 20px 24px;
        border: none;
        border-radius: 10px;
        color: #E0E0E0;
        background-color: transparent;
        margin: 6px 10px;
    }
    #SidebarButton:hover { 
        background-color: rgba(0, 123, 255, 0.15);
    }
    #SidebarButton:checked { 
        background-color: #007bff;
        color: white;
        font-weight: bold;
    }

    /* --- Dashboard Typography --- */
    #DashboardWelcomeTitle { 
        font-family: "Audiowide"; 
        font-size: 56px;
        color: #ffffff;
        margin-bottom: 10px;
        text-shadow: 0 0 15px rgba(0, 123, 255, 0.7);
    }
    #DashboardWelcomeSubtitle { 
        font-size: 22px;
        color: #b0b0b0;
        margin-bottom: 30px;
        text-shadow: 0 0 8px rgba(0, 123, 255, 0.4);
        font-style: italic;  /* Make subtitle italic */
    }
    #TitleLabel { 
        font-size: 34px;
        font-weight: bold;
        color: #ffffff;
        text-shadow: 0 0 12px rgba(0, 123, 255, 0.6);
        margin-bottom: 25px;
    }

    /* --- Glassmorphism Session Box --- */
    #SessionBox {
        background: rgba(20, 20, 30, 0.6);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
        padding: 25px;
        margin: 10px 0;
        min-width: 500px;
        max-width: 550px;
    }
    #SessionBox[class="glass-box"] {
        background: rgba(20, 20, 30, 0.5);
        backdrop-filter: blur(15px);
        -webkit-backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.25);
        box-shadow: 0 8px 32px rgba(0, 123, 255, 0.15);
    }

    /* --- Modern Dropdowns --- */
    QComboBox#InnerComboBox {
        background: rgba(30, 30, 45, 0.7);
        backdrop-filter: blur(5px);
        -webkit-backdrop-filter: blur(5px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 15px;
        padding: 18px 22px;
        margin: 0 0 15px 0;
        min-height: 65px;
        font-size: 28px;
        color: #ffffff;
    }
    QComboBox#InnerComboBox:focus {
        border: 2px solid rgba(0, 123, 255, 0.7);
        box-shadow: 0 0 15px rgba(0, 123, 255, 0.3);
    }
    QComboBox#InnerComboBox QLineEdit {
        color: #ffffff;
        font-size: 30px;
        font-weight: 500;
        background: transparent;
    }
    QComboBox#InnerComboBox QLineEdit::placeholder {
        color: #aaaaaa;
        font-size: 30px;
        font-style: italic;
    }
    QComboBox#InnerComboBox QAbstractItemView {
        font-size: 24px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        selection-background-color: rgba(0, 123, 255, 0.5);
        background-color: rgba(30, 30, 45, 0.9);
        color: #ffffff;
        padding: 12px;
        border-radius: 10px;
    }
    QComboBox#InnerComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: center right;
        width: 50px;
        border: none;
        background: transparent;
    }
    QComboBox#InnerComboBox::down-arrow {
        image: url(images/dropdown_icon.png);
        width: 30px;
        height: 30px;
    }

    /* --- Modern Start Button --- */
    #StartButton {
        font-size: 26px;
        font-weight: bold;
        padding: 18px 35px;
        border-radius: 15px;
        border: none;
        background: rgba(60, 60, 80, 0.7);
        color: #888;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        margin-top: 20px;
    }
    #StartButton:!disabled {
        background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, 
            stop:0 rgba(0, 123, 255, 0.8), stop:1 rgba(0, 86, 179, 0.8));
        color: white;
        box-shadow: 0 5px 15px rgba(0, 123, 255, 0.4);
    }
    #StartButton:!disabled:hover {
        background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, 
            stop:0 rgba(0, 105, 217, 0.9), stop:1 rgba(0, 64, 133, 0.9));
        box-shadow: 0 8px 20px rgba(0, 123, 255, 0.6);
    }

    /* --- Manage Batch Page --- */
    #SettingsHeader { 
        font-size: 26px;
        color: #ffffff;
        text-shadow: 0 0 10px rgba(0, 123, 255, 0.5);
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 2px solid rgba(0, 123, 255, 0.3);
    }
    QListWidget {
        font-size: 22px;
        background: rgba(30, 30, 45, 0.7);
        backdrop-filter: blur(5px);
        -webkit-backdrop-filter: blur(5px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 12px;
        padding: 10px;
        color: #ffffff;
    }
    QListWidget::item:selected {
        background-color: rgba(0, 123, 255, 0.4);
        color: #ffffff;
    }
    #SettingsLineEdit {
        font-size: 22px;
        padding: 15px;
        background: rgba(30, 30, 45, 0.7);
        backdrop-filter: blur(5px);
        -webkit-backdrop-filter: blur(5px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 12px;
        color: #ffffff;
    }
    #AddButton, #DeleteButton {
        font-size: 22px;
        padding: 12px 20px;
        font-weight: bold;
        border-radius: 10px;
        border: none;
    }
    #AddButton {
        background: rgba(40, 167, 69, 0.8);
        color: white;
    }
    #AddButton:hover {
        background: rgba(33, 136, 56, 0.9);
    }
    #DeleteButton {
        background: rgba(220, 53, 69, 0.8);
        color: white;
    }
    #DeleteButton:hover {
        background: rgba(200, 35, 51, 0.9);
    }
"""
# --- (ORIGINAL FUNCTIONS - 100% UNTOUCHED) ---
VALID_COLUMN_NAMES = {
    "image number": "A", "purohit name": "B", "bahi name": "C", "bahi number": "D", "folio number": "E",
    "district": "G", "tehsil": "H", "station": "I", "post office": "J", "pincode": "K",
    "village": "L", "city": "L", "place": "M", "kuldevta": "N", "gotra": "O", "caste": "P",
    "sub caste": "Q", "individual id": "R", "given name": "S", "first name": "S", "surname": "T",
    "last name": "T", "relation": "U", "family id": "V", "gender": "W", "ritual date": "X",
    "ritual name one": "Y", "whose ritual one": "Z", "ritual name two": "AA", "whose ritual two": "AB",
    "contact number one": "AC", "contact number two": "AD", "country id": "AE", "state id": "AF",
    "district id": "AG", "village id": "AH", "city id": "AH", "additional field one": "AI",
    "additional field two": "AJ", "additional field three": "AK", "additional field four": "AL",
    "remarks": "AM", "flag": "AM"
}
COMMON_SURNAMES = [
    "अग्रवाल", "गुप्ता", "शर्मा", "वर्मा", "अग्निहोत्री", "तिवारी", "पांडे", "द्विवेदी",
    "चतुर्वेदी", "त्रिवेदी", "पाठक", "शुक्ला", "वाजपेयी", "भट्ट", "मिश्रा", "दुबे",
    "तिवारी", "पांडेय", "जोशी", "रावत", "नेगी", "पंत", "बहुगुणा", "डिमरी", "सक्सेना",
    "वर्मा", "मालवीय", "चौहान", "सोलंकी", "परमार", "राठौड़", "सिसोदिया", "चंदेल",
    "तोमर", "हड्डा", "बुंदेला", "पटेल", "देशमुख", "पाटिल", "रेड्डी", "नायर", "चेट्टियार",
    "गांधी", "नेहरू", "शर्मा", "वर्मा", "गुप्ता", "अग्रवल", "जैन", "सिंह",
    "कुमार", "कुमारी", "यादव", "चौधरी", "गुर्जर", "जाट", "अहीर", "यादव", "सैनी",
    "ठाकुर", "राजपूत", "शेखावत", "कुशवाहा", "मौर्य", "गुप्ता", "मौर्य"
]

def words_to_numbers(text):
    num_map = {'शून्य': '0', 'एक': '1', 'दो': '2', 'तीन': '3', 'चार': '4', 'पांच': '5', 'पाँच': '5', 'छह': '6',
               'सात': '7', 'आठ': '8', 'नौ': '9', 'दस': '10'}
    for word, digit in num_map.items():
        text = text.replace(word, digit)
    return text


def convert_hindi_date(hindi_date_str):
    hindi_to_english_num = {'एक': '1', 'दो': '2', 'तीन': '3', 'चार': '4', 'पांच': '5', 'पाँच': '5', 'छह': '6',
                            'सात': '7', 'आठ': '8', 'नौ': '9', 'दस': '10', 'ग्यारह': '11', 'बारह': '12', 'तेरह': '13',
                            'चौदह': '14', 'पंद्रह': '15', 'सोलह': '16', 'सत्रह': '17', 'अठारह': '18', 'उन्नीस': '19',
                            'बीस': '20', 'इक्कीस': '21', 'बाईस': '22', 'तेईस': '23', 'चौबीस': '24', 'पच्चीस': '25',
                            'छब्बीस': '26', 'सत्ताईस': '27', 'अट्ठाईस': '28', 'उनतीस': '29', 'तीस': '30',
                            'इकत्तीस': '31'}
    months = ['चैत', 'चैत्र', 'बैसाख', 'वैशाख', 'जेठ', 'ज्येष्ठ', 'आषाढ़', 'सावन', 'श्रावण', 'भादो', 'भाद्रपद', 'क्वार',
              'आश्विन', 'कार्तिक', 'अगहन', 'मार्गशीर्ष', 'पूस', 'पौष', 'माघ', 'फाल्गुन']
    words = hindi_date_str.split()
    day_digit, month_name, year_digit = "", "", ""
    year_parts_started = False
    for word in words:
        if word in hindi_to_english_num and not year_parts_started:
            day_digit = hindi_to_english_num[word]
        elif word in months:
            month_name = word
            year_parts_started = True
        elif year_parts_started:
            year_digit += hindi_to_english_num.get(word, word)
    if "हजार" in year_digit or "सौ" in year_digit:
        year_digit = year_digit.replace("दो हजार", "20").replace("दो हज़ार", "20")
    if day_digit and month_name and year_digit:
        return f"{day_digit} {month_name} {year_digit}"
    else:
        return hindi_date_str


def extract_surname(full_name):
    if not full_name: return "", full_name
    words = str(full_name).split()
    if len(words) < 2: return "", full_name
    last_word = words[-1]
    if last_word in COMMON_SURNAMES:
        return last_word, " ".join(words[:-1])
    else:
        return "", full_name


# --- (YOUR ORIGINAL ControlBox CLASS - 100% UNTOUCHED) ---
class ControlBox(QWidget):
    start_clicked = Signal()
    stop_clicked = Signal()
    minimize_app = Signal()

    def __init__(self, parent=None):
        super().__init__()
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint | Qt.WindowMinimizeButtonHint)
        self.setFixedSize(380, 200)
        self.setStyleSheet("""
            background-color: #2b2b2b; color: white; border-radius: 12px;
            border: 1px solid #444; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3);
        """)
        self._mouse_pos = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 12, 15, 12)
        layout.setSpacing(8)
        title_row = QHBoxLayout()
        title = QLabel("Speakify Voice Control")
        title.setStyleSheet("font-weight: bold; font-size: 16px; color: #4CAF50;")
        title_row.addWidget(title)
        title_row.addStretch()
        self.min_button = QPushButton("—")
        self.min_button.setFixedSize(24, 24)
        self.min_button.clicked.connect(self.showMinimized)
        self.min_button.setStyleSheet("""
            QPushButton { background-color: transparent; color: #ccc; border: none; font-weight: bold; font-size: 14px; }
            QPushButton:hover { color: white; background-color: #444; border-radius: 4px; }
        """)
        title_row.addWidget(self.min_button)
        self.min_app_button = QPushButton("Minimize App")
        self.min_app_button.setFixedSize(100, 24)
        self.min_app_button.clicked.connect(self.minimize_app.emit)
        self.min_app_button.setStyleSheet("""
            QPushButton { background-color: #555; color: white; border: none; border-radius: 4px; font-size: 10px; padding: 2px; }
            QPushButton:hover { background-color: #666; }
        """)
        title_row.addWidget(self.min_app_button)
        layout.addLayout(title_row)
        self.status_label = QLabel("Active Column: None\nNext Row: -    ID: -")
        self.status_label.setStyleSheet("font-size: 13px; padding: 8px; background-color: #333; border-radius: 6px;")
        layout.addWidget(self.status_label)
        self.recording_label = QLabel("")
        self.recording_label.setStyleSheet("font-size: 14px; padding: 8px; color: #4CAF50; font-weight: bold;")
        self.recording_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.recording_label)
        btn_row = QHBoxLayout()
        self.start_button = QPushButton("Start")
        self.stop_button = QPushButton("Stop")
        self.stop_button.setEnabled(False)
        button_style = """
            QPushButton { font-weight: bold; padding: 8px 15px; border-radius: 6px; border: none; }
        """
        self.start_button.setStyleSheet(
            button_style + "QPushButton { background-color: #4CAF50; } QPushButton:hover { background-color: #45a049; }")
        self.stop_button.setStyleSheet(
            button_style + "QPushButton { background-color: #f44336; } QPushButton:hover { background-color: #d32f2f; }")
        btn_row.addWidget(self.start_button)
        btn_row.addWidget(self.stop_button)
        layout.addLayout(btn_row)
        self.helper_label = QLabel("Click on any Excel column to activate it for voice entry")
        self.helper_label.setStyleSheet("font-size: 11px; color: #cfcfcf; padding-top: 5px;")
        layout.addWidget(self.helper_label)
        self.timer = QTimer(self)
        self.timer.setInterval(350)
        self._anim_state = 0
        self.timer.timeout.connect(self._update_animation)
        self.start_button.clicked.connect(self._on_start)
        self.stop_button.clicked.connect(self._on_stop)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton: self._mouse_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft(); event.accept()

    def mouseMoveEvent(self, event):
        if self._mouse_pos is not None and event.buttons() & Qt.LeftButton: self.move(
            event.globalPosition().toPoint() - self._mouse_pos); event.accept()

    def mouseReleaseEvent(self, event):
        self._mouse_pos = None;
        event.accept()

    def _on_start(self):
        self.start_button.setEnabled(False);
        self.stop_button.setEnabled(
            True);
        self.timer.start();
        self.start_clicked.emit()

    def _on_stop(self):
        self.start_button.setEnabled(True);
        self.stop_button.setEnabled(
            False);
        self.timer.stop();
        self.recording_label.setText("");
        self.stop_clicked.emit()

    def _update_animation(self):
        self.recording_label.setText(f"Listening{'.' * ((self._anim_state % 3) + 1)}");
        self._anim_state += 1

    def set_active_column(self, col_name, next_row="-", individual_id="-"):
        self.status_label.setText(f"Active Column: {col_name}\nNext Row: {next_row}    ID: {individual_id}")

    def update_status(self, text, is_listening=False):
        try:
            if is_listening:
                self.recording_label.setText(text)
                if not self.timer.isActive(): self.timer.start(); self.start_button.setEnabled(
                    False); self.stop_button.setEnabled(True)
            else:
                self.status_label.setText(text + "\n" + "\n".join(self.status_label.text().splitlines()[1:]))
                if self.timer.isActive(): self.timer.stop()
                self.recording_label.setText("");
                self.start_button.setEnabled(True);
                self.stop_button.setEnabled(False)
        except Exception:
            pass


# --- (YOUR ORIGINAL VoiceWorker CLASS - 100% UNTOUCHED) ---
class VoiceWorker(QObject):
    update_status_signal = Signal(str, bool)
    recognized_text_signal = Signal(str)

    def __init__(self, app_instance):
        super().__init__()
        self.app = app_instance
        self.recognizer = sr.Recognizer()
        self.recognizer.pause_threshold = 0.6
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.energy_threshold = 4000
        self.is_running = False

    def run(self):
        self.is_running = True
        while self.is_running:
            language_code = 'hi-IN'
            try:
                with sr.Microphone() as source:
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                    prompt = self.app.current_column_name if self.app.current_column else "No column selected"
                    self.update_status_signal.emit(f"Listening for Hindi data in '{prompt}' column", True)
                    audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=15)
                    self.update_status_signal.emit("Processing...", False)
                    try:
                        text = self.recognizer.recognize_google(audio, language=language_code, show_all=False)
                        self.recognized_text_signal.emit(text.lower().strip())
                    except sr.UnknownValueError:
                        print("Could not understand audio");
                        self.update_status_signal.emit(
                            "Could not understand Hindi audio", False)
                    except Exception as e:
                        print("Recognition error:", e);
                        self.update_status_signal.emit("Recognition error", False)
            except sr.WaitTimeoutError:
                print("Listening timeout...")
            except Exception as e:
                print("Microphone/listen error:", e);
                self.update_status_signal.emit("Mic error", False)
        try:
            self.update_status_signal.emit("Stopped", False)
        except Exception:
            pass

    def stop(self):
        self.is_running = False


# --- Main Application ---
class SpeakifyApp(QMainWindow):
    session_started_signal = Signal(str)
    column_selected_signal = Signal(str)
    excel_disconnected_signal = Signal()

    def __init__(self):
        super().__init__()
        self.setObjectName("MainWindow")
        self.setWindowTitle("Speakify v2.0 - Final Edition")
        self.setGeometry(100, 100, 1280, 800)
        self.setStyleSheet(STYLESHEET)
        self.setWindowIcon(qta.icon('fa5s.microphone-alt', color='white'))

        self.config_file = "config.json"
        self.purohit_list = []
        self.bahi_list = []
        self.excel_files = [] # Ye line add ki gayi hai
        self.load_config()

        # Is section ko load_config ke neeche move kar diya gaya hai
        self.wb = None
        self.excel_app = None
        self.sheet = None
        self.excel_connected = False
        self.lock = threading.Lock()
        self.current_column = None
        self.current_column_name = ""
        self.last_data_position = ""
        self.individual_id_counter = 1
        self.selected_excel_file = ""
        self.selected_image_number = ""
        self.selected_purohit = ""
        self.selected_bahi = ""
        self.last_seen = {
            'gotra': None, 'caste': None, 'sub_caste': None,
            'village': None, 'place': None
        }
        self.filled_columns = set()
        self.voice_worker = None
        self.voice_thread = None
        self.recognized_text_display = QTextEdit()
        self.recognized_text_display.setReadOnly(True)
        self.active_column_label = QLabel("Active Column: None")
        self.active_column_label.setObjectName("ActiveColumnLabel")
        self.selection_timer = QTimer(self)
        self.selection_timer.timeout.connect(self.check_excel_selection)
        self.last_selection = ""
        self.connection_timer = QTimer(self)
        self.connection_timer.timeout.connect(self.check_excel_connection)
        self.connection_timer.start(5000)

        self.setup_ui() # UI setup hamesha end me

        self.session_started_signal.connect(self.update_live_screen_info)
        self.column_selected_signal.connect(self.on_column_selected)
        self.excel_disconnected_signal.connect(self.on_excel_disconnected)
        self.control_box.start_clicked.connect(self.start_listening)
        self.control_box.stop_clicked.connect(self.stop_listening)
        self.control_box.minimize_app.connect(self.minimize_app)

    def minimize_app(self):
        self.showMinimized()

    def load_config(self):
        try:
            # Excel files bhi yahin load kar lo
            if os.path.isdir(EXCEL_FILES_FOLDER_PATH):
                self.excel_files = [f for f in os.listdir(EXCEL_FILES_FOLDER_PATH) if f.endswith(('.xlsx', '.xls'))]

            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.purohit_list = config.get("purohits", [])
                    self.bahi_list = config.get("bahis", [])
            else:
                self.purohit_list = ["RAMJI SHARMA", "PANDIT1", "Shailendra Kumar"]
                self.bahi_list = ["BAHI1", "BAHI2", "KHATI RAJAN"]
                self.save_config()
        except Exception as e:
            print(f"Error loading config file: {e}")
            self.purohit_list = []
            self.bahi_list = []

    def add_new_purohit(self):
        new_name = self.new_purohit_input.text().strip().upper()
        if new_name and new_name not in self.purohit_list:
            self.purohit_list.append(new_name)
            self.save_config()
            # Dono jagah update karo
            self.purohit_list_widget.addItem(new_name)
            self.purohit_dropdown.addItem(new_name)
            self.new_purohit_input.clear()
            QMessageBox.information(self, "Success", f"'{new_name}' has been added.")
        else:
            QMessageBox.warning(self, "Error", "Purohit name cannot be empty or a duplicate.")

    def add_new_bahi(self):
        new_name = self.new_bahi_input.text().strip().upper()
        if new_name and new_name not in self.bahi_list:
            self.bahi_list.append(new_name)
            self.save_config()
            # Dono jagah update karo
            self.bahi_list_widget.addItem(new_name)
            self.bahi_dropdown.addItem(new_name)
            self.new_bahi_input.clear()
            QMessageBox.information(self, "Success", f"'{new_name}' has been added.")
        else:
            QMessageBox.warning(self, "Error", "Bahi name cannot be empty or a duplicate.")

    # Is function ki ab zaroorat nahi, to isko delete kar dein ya comment kar dein
    # def load_excel_files_into_dropdown(self):
    #     if not os.path.isdir(EXCEL_FILES_FOLDER_PATH): return
    #     files = [f for f in os.listdir(EXCEL_FILES_FOLDER_PATH) if f.endswith(('.xlsx', '.xls'))]
    #     self.excel_file_dropdown.addItems(files)

    # --- ✨ UI SETUP FUNCTION WITH ALL YOUR REQUESTS ---
    # --- ✨ UI SETUP FUNCTION WITH ALL YOUR REQUESTS ---
    def setup_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Background setup
        self.background_label = QLabel(central_widget)
        pixmap = QPixmap("images/blue_abstract_tech_background.jpg")
        self.background_label.setPixmap(pixmap)
        self.background_label.setScaledContents(True)

        # Main layout
        main_content_layout = QHBoxLayout(central_widget)
        main_content_layout.setContentsMargins(0, 0, 0, 0)
        main_content_layout.setSpacing(0)

        # Sidebar
        sidebar_frame = QFrame()
        sidebar_frame.setObjectName("Sidebar")
        sidebar_frame.setFixedWidth(280)  # Increased width
        sidebar_layout = QVBoxLayout(sidebar_frame)
        sidebar_layout.setAlignment(Qt.AlignTop)
        sidebar_layout.setContentsMargins(0, 20, 0, 20)
        sidebar_layout.setSpacing(10)

        # Logo placeholder in sidebar
        logo_placeholder = QFrame()
        logo_placeholder.setObjectName("LogoPlaceholder")
        logo_placeholder.setFixedHeight(80)
        sidebar_layout.addWidget(logo_placeholder)

        # Add "Speakify" text in sidebar
        sidebar_title = QLabel("Speakify")
        sidebar_title.setObjectName("SidebarTitle")
        sidebar_title.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(sidebar_title)

        # Navigation buttons with enhanced styling
        self.dashboard_button = QPushButton(qta.icon('fa5s.home', color='#E0E0E0'), "  Dashboard")
        self.live_entry_button = QPushButton(qta.icon('fa5s.microphone-alt', color='#E0E0E0'), "  Live Entry")
        self.manage_batch_button = QPushButton(qta.icon('fa5s.users-cog', color='#E0E0E0'), "  Manage Batch")

        buttons = [self.dashboard_button, self.live_entry_button, self.manage_batch_button]
        for btn in buttons:
            btn.setObjectName("SidebarButton")
            btn.setIconSize(QSize(24, 24))  # Increased icon size
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()

        # Button group
        for btn in buttons:
            btn.setCheckable(True)
        self.sidebar_button_group = QButtonGroup(self)
        self.sidebar_button_group.setExclusive(True)
        for btn in buttons:
            self.sidebar_button_group.addButton(btn)
        self.dashboard_button.setChecked(True)

        # Stacked widget for different screens
        self.stacked_widget = QStackedWidget()
        self.dashboard_screen = self.create_dashboard_screen()
        self.live_entry_screen = self.create_live_entry_screen()
        self.manage_batch_screen = self.create_settings_screen()

        self.stacked_widget.addWidget(self.dashboard_screen)
        self.stacked_widget.addWidget(self.live_entry_screen)
        self.stacked_widget.addWidget(self.manage_batch_screen)

        # Connect buttons to switch screens
        self.dashboard_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.dashboard_screen))
        self.live_entry_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.live_entry_screen))
        self.manage_batch_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.manage_batch_screen))

        # Add widgets to main layout
        main_content_layout.addWidget(sidebar_frame)
        main_content_layout.addWidget(self.stacked_widget)

        # Control box
        self.control_box = ControlBox(self)
        self.control_box.hide()
    # --- ✨ REDESIGNED DASHBOARD SCREEN ---
    # --- ✨ REDESIGNED DASHBOARD SCREEN ---
    def create_dashboard_screen(self):
        dashboard_widget = QWidget()
        dashboard_widget.setObjectName("DashboardWidget")

        # Main layout with background
        main_layout = QVBoxLayout(dashboard_widget)
        main_layout.setContentsMargins(50, 50, 50, 50)
        main_layout.setSpacing(30)

        # Header section with logo and title
        header_layout = QHBoxLayout()
        header_layout.setSpacing(50)

        # Left side - Logo and app name
        left_header = QVBoxLayout()
        left_header.setAlignment(Qt.AlignCenter)

        logo_label = QLabel()
        logo_pixmap = QPixmap("images/logo_graphic.png")
        logo_label.setPixmap(logo_pixmap.scaled(280, 280, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignCenter)

        # Add logo to layout
        left_header.addWidget(logo_label)

        # Add italic subtitle below logo
        app_subtitle = QLabel("Voice-Powered Data Entry Assistant")
        app_subtitle.setObjectName("DashboardWelcomeSubtitle")
        app_subtitle.setAlignment(Qt.AlignCenter)
        app_subtitle.setStyleSheet("font-style: italic;")  # Make it italic
        left_header.addWidget(app_subtitle)

        # Right side - Session creation box
        right_header = QVBoxLayout()
        right_header.setAlignment(Qt.AlignCenter)

        # Session box with glassmorphism effect
        session_box = QFrame()
        session_box.setObjectName("SessionBox")
        session_box.setProperty("class", "glass-box")  # Add property for glass effect

        session_layout = QVBoxLayout(session_box)
        session_layout.setContentsMargins(25, 25, 25, 25)
        session_layout.setSpacing(20)

        # Title - Changed to "Session Setup" and left-aligned
        title = QLabel("Session Setup")
        title.setObjectName("TitleLabel")
        title.setAlignment(Qt.AlignLeft)
        session_layout.addWidget(title)

        # Dropdowns with enhanced styling
        self.purohit_dropdown = QComboBox()
        self.purohit_dropdown.setObjectName("InnerComboBox")
        self.purohit_dropdown.addItems(self.purohit_list)
        self.purohit_dropdown.setEditable(True)
        self.purohit_dropdown.lineEdit().setPlaceholderText("Select Purohit")
        self.purohit_dropdown.setCurrentIndex(-1)
        purohit_completer = QCompleter(self.purohit_list)
        purohit_completer.setFilterMode(Qt.MatchContains)
        purohit_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.purohit_dropdown.setCompleter(purohit_completer)
        self.purohit_dropdown.currentTextChanged.connect(self.check_selection)
        self.purohit_dropdown.activated.connect(self.check_selection)

        self.bahi_dropdown = QComboBox()
        self.bahi_dropdown.setObjectName("InnerComboBox")
        self.bahi_dropdown.addItems(self.bahi_list)
        self.bahi_dropdown.setEditable(True)
        self.bahi_dropdown.lineEdit().setPlaceholderText("Select Bahi")
        self.bahi_dropdown.setCurrentIndex(-1)
        bahi_completer = QCompleter(self.bahi_list)
        bahi_completer.setFilterMode(Qt.MatchContains)
        bahi_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.bahi_dropdown.setCompleter(bahi_completer)
        self.bahi_dropdown.currentTextChanged.connect(self.check_selection)
        self.bahi_dropdown.activated.connect(self.check_selection)

        self.excel_file_dropdown = QComboBox()
        self.excel_file_dropdown.setObjectName("InnerComboBox")
        self.excel_file_dropdown.addItems(self.excel_files)
        self.excel_file_dropdown.setEditable(True)
        self.excel_file_dropdown.lineEdit().setPlaceholderText("Select Excel File")
        self.excel_file_dropdown.setCurrentIndex(-1)
        excel_completer = QCompleter(self.excel_files)
        excel_completer.setFilterMode(Qt.MatchContains)
        excel_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.excel_file_dropdown.setCompleter(excel_completer)
        self.excel_file_dropdown.currentTextChanged.connect(self.check_selection)
        self.excel_file_dropdown.activated.connect(self.check_selection)

        # Add dropdowns to session box
        session_layout.addWidget(self.purohit_dropdown)
        session_layout.addWidget(self.bahi_dropdown)
        session_layout.addWidget(self.excel_file_dropdown)

        # Start button centered
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.start_button = QPushButton("Start Session")
        self.start_button.setObjectName("StartButton")
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.start_session)
        button_layout.addWidget(self.start_button)
        button_layout.addStretch()
        session_layout.addLayout(button_layout)

        right_header.addWidget(session_box)
        right_header.addStretch()

        # Add both sides to header
        header_layout.addLayout(left_header, 1)
        header_layout.addLayout(right_header, 1)

        # Add header to main layout
        main_layout.addLayout(header_layout)
        main_layout.addStretch()

        return dashboard_widget

    def create_settings_screen(self):
        settings_widget = QWidget()
        settings_widget.setObjectName("SettingsWidget")

        main_layout = QVBoxLayout(settings_widget)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(30)

        # Title
        title = QLabel("Manage Batch Lists")
        title.setObjectName("TitleLabel")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # Content layout with two columns
        content_layout = QHBoxLayout()
        content_layout.setSpacing(40)

        # Purohit Management Section
        purohit_frame = QFrame()
        purohit_frame.setObjectName("SettingsFrame")
        purohit_layout = QVBoxLayout(purohit_frame)

        purohit_header = QLabel("Purohit Names")
        purohit_header.setObjectName("SettingsHeader")
        purohit_layout.addWidget(purohit_header)

        self.purohit_list_widget = QListWidget()
        self.purohit_list_widget.addItems(self.purohit_list)
        purohit_layout.addWidget(self.purohit_list_widget)

        add_purohit_layout = QHBoxLayout()
        self.new_purohit_input = QLineEdit()
        self.new_purohit_input.setObjectName("SearchBar")
        self.new_purohit_input.setPlaceholderText("Enter new Purohit name...")
        add_purohit_btn = QPushButton("Add")
        add_purohit_btn.setObjectName("AddButton")
        add_purohit_btn.clicked.connect(self.add_new_purohit)
        add_purohit_layout.addWidget(self.new_purohit_input)
        add_purohit_layout.addWidget(add_purohit_btn)
        purohit_layout.addLayout(add_purohit_layout)

        delete_purohit_btn = QPushButton("Delete Selected")
        delete_purohit_btn.setObjectName("DeleteButton")
        delete_purohit_btn.clicked.connect(self.delete_selected_purohit)
        purohit_layout.addWidget(delete_purohit_btn)

        # Bahi Management Section
        bahi_frame = QFrame()
        bahi_frame.setObjectName("SettingsFrame")
        bahi_layout = QVBoxLayout(bahi_frame)

        bahi_header = QLabel("Bahi Names")
        bahi_header.setObjectName("SettingsHeader")
        bahi_layout.addWidget(bahi_header)

        self.bahi_list_widget = QListWidget()
        self.bahi_list_widget.addItems(self.bahi_list)
        bahi_layout.addWidget(self.bahi_list_widget)

        add_bahi_layout = QHBoxLayout()
        self.new_bahi_input = QLineEdit()
        self.new_bahi_input.setObjectName("SearchBar")
        self.new_bahi_input.setPlaceholderText("Enter new Bahi name...")
        add_bahi_btn = QPushButton("Add")
        add_bahi_btn.setObjectName("AddButton")
        add_bahi_btn.clicked.connect(self.add_new_bahi)
        add_bahi_layout.addWidget(self.new_bahi_input)
        add_bahi_layout.addWidget(add_bahi_btn)
        bahi_layout.addLayout(add_bahi_layout)

        delete_bahi_btn = QPushButton("Delete Selected")
        delete_bahi_btn.setObjectName("DeleteButton")
        delete_bahi_btn.clicked.connect(self.delete_selected_bahi)
        bahi_layout.addWidget(delete_bahi_btn)

        content_layout.addWidget(purohit_frame)
        content_layout.addWidget(bahi_frame)

        main_layout.addLayout(content_layout)

        return settings_widget

    def delete_selected_purohit(self):
        selected_items = self.purohit_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Please select a Purohit name to delete.")
            return

        reply = QMessageBox.question(self, 'Confirm Delete',
                                     f"Are you sure you want to delete '{selected_items[0].text()}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            name_to_delete = selected_items[0].text()
            self.purohit_list.remove(name_to_delete)
            self.purohit_list_widget.takeItem(self.purohit_list_widget.row(selected_items[0]))

            # Dashboard dropdown se bhi update karo
            index = self.purohit_dropdown.findText(name_to_delete)
            if index >= 0:
                self.purohit_dropdown.removeItem(index)

            self.save_config()
            QMessageBox.information(self, "Success", f"'{name_to_delete}' has been deleted.")

    def delete_selected_bahi(self):
        selected_items = self.bahi_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Please select a Bahi name to delete.")
            return

        reply = QMessageBox.question(self, 'Confirm Delete',
                                     f"Are you sure you want to delete '{selected_items[0].text()}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            name_to_delete = selected_items[0].text()
            self.bahi_list.remove(name_to_delete)
            self.bahi_list_widget.takeItem(self.bahi_list_widget.row(selected_items[0]))

            index = self.bahi_dropdown.findText(name_to_delete)
            if index >= 0:
                self.bahi_dropdown.removeItem(index)

            self.save_config()
            QMessageBox.information(self, "Success", f"'{name_to_delete}' has been deleted.")

    # --- (YOUR ORIGINAL FUNCTIONS START HERE - 100% UNTOUCHED) ---
    def create_live_entry_screen(self):
        live_widget = QWidget()
        layout = QVBoxLayout(live_widget)
        self.live_status_label = QLabel("Session not started yet.")
        self.live_status_label.setObjectName("TitleLabel")
        layout.addWidget(self.live_status_label, alignment=Qt.AlignCenter)
        status_row = QHBoxLayout()
        status_row.addWidget(self.active_column_label)
        status_row.addStretch()
        layout.addLayout(status_row)
        layout.addWidget(QLabel("Recognized Text:"))
        layout.addWidget(self.recognized_text_display)
        return live_widget

    def load_excel_files_into_dropdown(self):
        if not os.path.isdir(EXCEL_FILES_FOLDER_PATH): return
        files = [f for f in os.listdir(EXCEL_FILES_FOLDER_PATH) if f.endswith(('.xlsx', '.xls'))]
        self.excel_file_dropdown.addItems(files)

    def check_selection(self):
        if (
                self.purohit_dropdown.currentIndex() > 0 and
                self.bahi_dropdown.currentIndex() > 0 and
                self.excel_file_dropdown.currentIndex() > 0):
            self.start_button.setEnabled(True)
        else:
            self.start_button.setEnabled(False)

    def start_session(self):
        selected_file = self.excel_file_dropdown.currentText()
        purohit_name = self.purohit_dropdown.currentText()
        bahi_name = self.bahi_dropdown.currentText()
        full_path = os.path.join(EXCEL_FILES_FOLDER_PATH, selected_file)

        # Reset state variables before starting a new session
        self.reset_state_variables()

        try:
            self.selected_excel_file = selected_file
            self.selected_image_number = os.path.splitext(selected_file)[0]
            self.selected_purohit = purohit_name
            self.selected_bahi = bahi_name
            self.excel_app = xw.App(visible=True)
            self.wb = self.excel_app.books.open(full_path)
            if len(self.excel_app.books) > 1 and 'book1' in [b.name.lower() for b in self.excel_app.books]:
                try:
                    self.excel_app.books['Book1'].close()
                except Exception:
                    pass
            self.wb.activate()
            self.sheet = self.wb.sheets[0]
            self.excel_connected = True
            self.live_status_label.setText(
                "Session started. Click on any column in Excel to activate it for voice entry.")
            self.load_excel_state()
            self.refresh_last_seen_fields()
            self.stacked_widget.setCurrentWidget(self.live_entry_screen)
            session_info = f"Working on: {selected_file}\n{purohit_name} | {bahi_name}"
            self.session_started_signal.emit(session_info)
            self.control_box.show()
            try:
                self.control_box.move(self.geometry().topRight() - QPoint(self.control_box.width() + 20, -40))
            except Exception:
                pass
            try:
                last_row = self.sheet.range(f"A{self.sheet.cells.last_cell.row}").end('up').row
                next_row = last_row + 1 if last_row >= 1 else 2
            except Exception:
                next_row = 2
            self.control_box.set_active_column("None", next_row, self.individual_id_counter)
            self.selection_timer.start(500)
            self.minimize_app()
        except Exception as e:
            QMessageBox.critical(self, "Excel Error", f"Could not open the file.\n\nError: {e}")
            self.excel_connected = False

    def reset_state_variables(self):
        """Reset all state variables when starting a new session"""
        self.excel_files = []
        self.wb = None
        self.excel_app = None
        self.sheet = None
        self.excel_connected = False
        self.current_column = None
        self.current_column_name = ""
        self.last_data_position = ""
        self.individual_id_counter = 1
        self.selected_excel_file = ""
        self.selected_image_number = ""
        self.selected_purohit = ""
        self.selected_bahi = ""
        # FIXED: Initialize village and place separately
        self.last_seen = {
            'gotra': None,
            'caste': None,
            'sub_caste': None,
            'village': None,
            'place': None
        }
        self.filled_columns = set()
        self.last_selection = ""
        self.voice_worker = None
        self.voice_thread = None

    def check_excel_connection(self):
        if not self.excel_connected: return
        try:
            if self.excel_app and self.wb:
                test = self.wb.name
                test2 = self.excel_app.hwnd
            else:
                self.excel_disconnected_signal.emit()
        except Exception as e:
            print(f"Excel connection check failed: {e}")
            self.excel_disconnected_signal.emit()

    def on_excel_disconnected(self):
        print("Excel disconnected. Stopping timers.")
        self.excel_connected = False
        self.selection_timer.stop()
        self.live_status_label.setText("Excel connection lost. Please restart the session.")
        self.control_box.hide()
        QMessageBox.warning(self, "Excel Disconnected",
                            "Excel has been closed or is no longer responding.\n\nPlease start a new session to continue.")

    def check_excel_selection(self):
        if not self.excel_connected: return
        try:
            if not self.wb or not self.excel_app:
                self.excel_disconnected_signal.emit(); return
            try:
                test = self.wb.name
            except Exception:
                self.excel_disconnected_signal.emit(); return

            current_selection = self.wb.selection.address
            if current_selection != self.last_selection:
                # --- NEW LOGIC ---
                # User has moved from the last cell. Check if they entered data there.
                if self.last_selection:
                    self.handle_manual_autofill(self.last_selection)

                # Original logic: Update UI for the NEWLY selected cell
                try:
                    top_left = current_selection.split(':')[0]
                    selected_col_letter = self.sheet.range(top_left).column_letter
                except Exception:
                    selected_col_letter = ''.join(
                        [c for c in current_selection.split(':')[0].replace('$', '') if c.isalpha()])

                self.column_selected_signal.emit(selected_col_letter)
                self.last_selection = current_selection
        except Exception as e:
            print(f"Error checking Excel selection: {e}")
            if "disconnected" in str(e) or "invoked has disconnected" in str(e):
                self.excel_disconnected_signal.emit()

    def handle_manual_autofill(self, address_string):
        """
        This function triggers AFTER you enter data in a cell and move away.
        It handles the auto-filling for the row you just edited.
        """
        try:
            edited_cell = self.sheet.range(address_string.split(':')[0])
            col_letter = edited_cell.column_letter
            row = edited_cell.row

            trigger_columns = ['O', 'P', 'Q']
            # Subcaste ko is list se alag kar diya hai
            fields_to_fill = {
                'L': 'village', 'M': 'place', 'O': 'gotra', 'P': 'caste'
            }

            if col_letter in trigger_columns and row > 1 and edited_cell.value is not None:
                print(f"Manual entry in {col_letter}{row} detected. Checking for auto-fill...")
                self.refresh_last_seen_fields()

                # Normal fields ko pehle ki tarah auto-fill karo
                for col, key in fields_to_fill.items():
                    if col != col_letter.upper():
                        target_cell = self.sheet.range(f"{col}{row}")
                        if target_cell.value is None and self.last_seen[key]:
                            target_cell.value = self.last_seen[key]
                            print(f"Auto-filled '{key}' in row {row}.")

                # --- SUBCASTE KE LIYE SPECIAL CHECK ---
                # Subcaste (Q) tabhi fill karo jab Data Position (F) khaali na ho
                subcaste_cell = self.sheet.range(f"Q{row}")
                if subcaste_cell.value is None and self.last_seen['sub_caste']:
                    if self.sheet.range(f"F{row}").value is not None:
                        subcaste_cell.value = self.last_seen['sub_caste']
                        print(f"SUCCESS: Auto-filled 'sub_caste' in row {row} because DP exists.")
                    else:
                        print(f"INFO: Skipped auto-filling 'sub_caste' in row {row} because DP is missing.")

        except Exception as e:
            print(f"Error during handle_manual_autofill: {e}")

    def on_column_selected(self, selected_col_letter):
        try:
            found_command = ""
            for command, col_letter in VALID_COLUMN_NAMES.items():
                if col_letter == selected_col_letter:
                    found_command = command
                    break
            with self.lock:
                self.current_column = selected_col_letter
                self.current_column_name = found_command
            try:
                last_row = self.sheet.range(f"{selected_col_letter}{self.sheet.cells.last_cell.row}").end('up').row
                next_row = last_row + 1 if last_row >= 1 else 2
            except Exception:
                next_row = 2

            # This function is now ONLY responsible for updating the UI labels
            if found_command:
                self.active_column_label.setText(
                    f"Selected Column: {found_command} ({selected_col_letter})\nReady for Hindi voice entry")
                self.control_box.set_active_column(found_command + f" ({selected_col_letter})", next_row,
                                                   self.individual_id_counter)
            else:
                self.active_column_label.setText(f"Selected Column: {selected_col_letter}\nReady for Hindi voice entry")
                self.control_box.set_active_column(selected_col_letter, next_row, self.individual_id_counter)
        except Exception as e:
            print(f"Error in on_column_selected: {e}")

    def start_listening(self):
        if not self.excel_connected:
            QMessageBox.warning(self, "Excel Not Connected",
                                "Please start a session and ensure Excel is open before using voice entry.")
            return
        try:
            if self.voice_worker is None:
                self.voice_worker = VoiceWorker(self)
                self.voice_worker.recognized_text_signal.connect(self.process_recognized_text)
                self.voice_worker.update_status_signal.connect(self.on_voice_status_update)
            if self.voice_thread is None or not getattr(self.voice_thread, "is_alive", lambda: False)():
                self.voice_thread = threading.Thread(target=self.voice_worker.run, daemon=True)
                self.voice_thread.start()
            else:
                self.voice_worker.is_running = True
            self.control_box.update_status("Listening...", True)
            print("Voice listening started.")
        except Exception as e:
            print("Error starting listening:", e)

    def stop_listening(self):
        try:
            if getattr(self, "voice_worker", None):
                try:
                    self.voice_worker.stop()
                except Exception as e:
                    print("Error calling voice_worker.stop():", e)
            try:
                if getattr(self, "voice_thread", None) and getattr(self.voice_thread, "is_alive",
                                                                   lambda: False)():
                    time.sleep(0.2)
            except Exception:
                pass
            self.voice_worker = None
            self.voice_thread = None
        except Exception as e:
            print("Error in stop_listening:", e)
        finally:
            try:
                self.control_box.update_status("Stopped", False)
            except Exception:
                pass
            print("Voice listening stopped.")

    def on_voice_status_update(self, text, is_listening):
        try:
            self.control_box.update_status(text, is_listening)
        except Exception:
            pass

    def refresh_dp_state(self):
        if not self.excel_connected: return
        try:
            if not self.sheet: return
            last_row_num = self.sheet.range(f"F{self.sheet.cells.last_cell.row}").end('up').row
            if last_row_num is None or last_row_num < 2: return
            dp_val = self.sheet.range(f"F{last_row_num}").value
            if dp_val is None: return
            dp_val = str(dp_val)
            if dp_val != self.last_data_position:
                self.last_data_position = dp_val
                last_id_for_dp = 0
                for r in range(last_row_num, 1, -1):
                    try:
                        if str(self.sheet.range(f"F{r}").value) == self.last_data_position:
                            rid = self.sheet.range(f"R{r}").value
                            if rid is not None:
                                last_id_for_dp = int(rid)
                                break
                    except Exception:
                        continue
                self.individual_id_counter = last_id_for_dp + 1 if last_id_for_dp >= 0 else 1
                print(f"[refresh_dp_state] DP='{self.last_data_position}', next ID='{self.individual_id_counter}'")
        except Exception as e:
            print("refresh_dp_state error:", e)

    def refresh_last_seen_fields(self):
        if not self.excel_connected: return
        try:
            if not self.sheet: return

            # Refresh gotra
            try:
                r = self.sheet.range(f"O{self.sheet.cells.last_cell.row}").end('up').row
                val = self.sheet.range(f"O{r}").value
                if val is not None and val != '':
                    self.last_seen['gotra'] = val
            except Exception:
                pass

            # Refresh caste
            try:
                r = self.sheet.range(f"P{self.sheet.cells.last_cell.row}").end('up').row
                val = self.sheet.range(f"P{r}").value
                if val is not None and val != '':
                    self.last_seen['caste'] = val
            except Exception:
                pass

            # Refresh sub_caste
            try:
                r = self.sheet.range(f"Q{self.sheet.cells.last_cell.row}").end('up').row
                val = self.sheet.range(f"Q{r}").value
                if val is not None and val != '':
                    self.last_seen['sub_caste'] = val
            except Exception:
                pass

            # Refresh village (column L) - FIXED: Only update village separately
            try:
                r = self.sheet.range(f"L{self.sheet.cells.last_cell.row}").end('up').row
                val = self.sheet.range(f"L{r}").value
                if val is not None and val != '':
                    self.last_seen['village'] = val
            except Exception:
                pass

            # Refresh place (column M) - FIXED: Only update place separately
            try:
                r = self.sheet.range(f"M{self.sheet.cells.last_cell.row}").end('up').row
                val = self.sheet.range(f"M{r}").value
                if val is not None and val != '':
                    self.last_seen['place'] = val
            except Exception:
                pass
        except Exception as e:
            print("refresh_last_seen_fields error:", e)

    def process_recognized_text(self, text):
        if not text or not self.excel_connected: return
        with self.lock:
            if self.current_column is None:
                self.recognized_text_display.append("Error: No column selected.")
                self.control_box.update_status("No column selected", False)
                return

            if self.sheet and self.current_column:
                try:
                    next_row = 2
                    while next_row <= 10000:
                        if self.sheet.range(f"{self.current_column}{next_row}").value is None: break
                        next_row += 1

                    final_data_to_write = convert_hindi_date(
                        text) if self.current_column.upper() == "X" else words_to_numbers(text)

                    is_given_name_column = self.current_column.upper() == "S"

                    if is_given_name_column:
                        self.filled_columns = set()
                        print("New person detected (Given Name entry). Resetting filled columns for this record.")

                    if is_given_name_column:
                        surname, remaining_name = extract_surname(final_data_to_write)
                        self.sheet.range(f'{self.current_column}{next_row}').value = remaining_name
                        if surname and not self.sheet.range(f"T{next_row}").value:
                            self.sheet.range(f"T{next_row}").value = surname
                    else:
                        self.sheet.range(f'{self.current_column}{next_row}').value = final_data_to_write

                    # --- FIX 3: भाई ko बाई me badalne wala code sahi jagah par ---
                    if is_given_name_column:
                        name_in_cell = str(self.sheet.range(f"S{next_row}").value or "")
                        if "भाई" in name_in_cell:
                            corrected_name = name_in_cell.replace("भाई", "बाई")
                            self.sheet.range(f"S{next_row}").value = corrected_name
                            print(f"Auto-corrected name from '{name_in_cell}' to '{corrected_name}'")
                    # --- BADLAV KHATM ---

                    if self.current_column.upper() == 'L':
                        self.last_seen['village'] = final_data_to_write
                    elif self.current_column.upper() == 'M':
                        self.last_seen['place'] = final_data_to_write

                    self.filled_columns.add(self.current_column)

                    if len(self.filled_columns) == 1:
                        self.sheet.range(f'A{next_row}').value = self.selected_image_number
                        self.sheet.range(f'B{next_row}').value = self.selected_purohit
                        self.sheet.range(f'C{next_row}').value = self.selected_bahi

                    try:
                        last_row_dp = self.sheet.range('F' + str(self.sheet.cells.last_cell.row)).end('up').row
                        dp_in_sheet = str(self.sheet.range(f'F{last_row_dp}').value)
                        if dp_in_sheet and dp_in_sheet != self.last_data_position:
                            self.last_data_position = dp_in_sheet
                            self.individual_id_counter = 1
                            self.last_seen['sub_caste'] = None
                            print(f"** Data Position changed to '{dp_in_sheet}'. SubCaste reset. **")
                    except:
                        pass

                    if self.last_data_position:
                        self.sheet.range(f'F{next_row}').value = self.last_data_position
                        if is_given_name_column:
                            self.sheet.range(f'R{next_row}').value = self.individual_id_counter
                            self.individual_id_counter += 1

                    # Auto-fill logic
                    if 'O' not in self.filled_columns and self.last_seen['gotra'] and not self.sheet.range(f"O{next_row}").value:
                        self.sheet.range(f"O{next_row}").value = self.last_seen['gotra']
                    if 'P' not in self.filled_columns and self.last_seen['caste'] and not self.sheet.range(f"P{next_row}").value:
                        self.sheet.range(f"P{next_row}").value = self.last_seen['caste']
                    if 'Q' not in self.filled_columns and self.last_seen['sub_caste'] and not self.sheet.range(f"Q{next_row}").value:
                        if self.sheet.range(f"F{next_row}").value is not None:
                             self.sheet.range(f"Q{next_row}").value = self.last_seen['sub_caste']
                    if 'L' not in self.filled_columns and self.last_seen['village'] and not self.sheet.range(f"L{next_row}").value:
                        self.sheet.range(f"L{next_row}").value = self.last_seen['village']
                    if 'M' not in self.filled_columns and self.last_seen['place'] and not self.sheet.range(f"M{next_row}").value:
                        self.sheet.range(f"M{next_row}").value = self.last_seen['place']

                    # --- FIX 4: Sahi Gender Logic ---
                    given_name = str(self.sheet.range(f"S{next_row}").value or "").lower()
                    relation = str(self.sheet.range(f"U{next_row}").value or "").lower()
                    if "की" in relation or any(s in given_name for s in ["बाई", "वाई", "राणी", "श्रीमती", "देवी"]):
                        gender = "महिला"
                    else:
                        gender = "पुरुष"
                    self.sheet.range(f"W{next_row}").value = gender
                    # --- BADLAV KHATM ---

                    self.refresh_last_seen_fields()
                    self.control_box.set_active_column(self.current_column_name + f" ({self.current_column})",
                                                       next_row + 1, self.individual_id_counter)
                except Exception as e:
                    print(f"Error in process_recognized_text: {e}")
                    self.recognized_text_display.append(f"Error processing data: {str(e)}")
                    if "disconnected" in str(e) or "invoked has disconnected" in str(e):
                        self.excel_disconnected_signal.emit()

    def load_excel_state(self):
        if not self.excel_connected: return
        self.last_data_position = ""
        self.individual_id_counter = 1
        try:
            last_row_num = self.sheet.range('F' + str(self.sheet.cells.last_cell.row)).end('up').row
            if last_row_num > 1:
                dp_val = self.sheet.range(f'F{last_row_num}').value
                if dp_val is not None: self.last_data_position = str(dp_val)
                last_id_for_dp = 0
                for i in range(last_row_num, 1, -1):
                    try:
                        if str(self.sheet.range(f'F{i}').value) == self.last_data_position:
                            rid = self.sheet.range(f'R{i}').value
                            if rid is not None:
                                last_id_for_dp = int(rid)
                                break
                    except Exception:
                        continue
                self.individual_id_counter = last_id_for_dp + 1 if last_id_for_dp >= 0 else 1
                print(f"Last state loaded: DP='{self.last_data_position}', Next ID='{self.individual_id_counter}'")
        except Exception as e:
            print(f"Nayi sheet lag rahi hai. Error: {e}")

    def update_live_screen_info(self, text):
        try:
            self.live_status_label.setText(text)
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            self.selection_timer.stop()
            self.connection_timer.stop()
            self.stop_listening()
            if self.wb: self.wb.close()
            if self.excel_app: self.excel_app.quit()
        except Exception as e:
            print(f"Error during cleanup: {e}")
        event.accept()

    # NEW: Add resizeEvent to handle background scaling
    def resizeEvent(self, event):
        self.background_label.setGeometry(0, 0, self.width(), self.height())
        super().resizeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # --- ✨ NEW: FONT & ANIMATION SETUP ---
    font_id = QFontDatabase.addApplicationFont("fonts/Audiowide-Regular.ttf")
    if font_id == -1: print("Warning: Custom font 'Audiowide-Regular.ttf' not loaded.")

    window = SpeakifyApp()

    window.setWindowOpacity(0.0)
    animation = QPropertyAnimation(window, b"windowOpacity", window)
    animation.setDuration(500)
    animation.setStartValue(0.0)
    animation.setEndValue(1.0)
    animation.setEasingCurve(QEasingCurve.InOutQuad)

    window.show()
    animation.start()
    sys.exit(app.exec())
