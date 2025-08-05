import sys
import subprocess
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QStatusBar, QLabel, QProgressBar, QMessageBox, QScrollArea,
    QSizePolicy, QPlainTextEdit, QHBoxLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSize
from PyQt6.QtGui import (QPalette, QLinearGradient, QColor, QBrush,
                        QFont, QPixmap, QIcon, QFontDatabase, QMovie)

# --- Константы и пути ---
SCRIPTS = {
    "ИНТЕРФАКС Бизнесс": os.path.join("parsers", "INTERFAX_Business_news.py"),
    "MASH Новости": os.path.join("parsers", "MASH_First_100_news.py"),
    "РИА Новости": os.path.join("parsers", "RIA_Ekonomika_news.py"),
    "ПРАЙМ Новости": os.path.join("parsers", "PRIME_news.py"),
    "RGru Новости": os.path.join("parsers", "RGru_news.py"),
    "ТАСС Новости": os.path.join("parsers", "TASS_news.py"),
    "ИНТЕРФАКС Новости": os.path.join("parsers", "INTERFAX_First_100_news.py"),
    "----": os.path.join("parsers", "Test.py")
}

class ParserWorker(QThread):
    """Класс для выполнения парсинга в отдельном потоке"""
    progress = pyqtSignal(int)  # Сигнал прогресса для UI
    finished = pyqtSignal(str, bool)  # Сигнал завершения (сообщение, успех)
    error = pyqtSignal(str)  # Сигнал ошибки
    console_output = pyqtSignal(str)  # Сигнал вывода в консоль

    def __init__(self, script_path):
        super().__init__()
        self.script_path = script_path
        self._is_running = True
        self.current_progress = 0
        self.last_progress = 0
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.smooth_progress)

    def run(self):
        """Основной метод выполнения парсера"""
        try:
            if not os.path.exists(self.script_path):
                raise FileNotFoundError(f"Файл скрипта не найден: {self.script_path}")

            self.console_output.emit(f"Запускаем парсер: {os.path.basename(self.script_path)}")
            self.console_output.emit(f"Прогружаем страницу сайта, находим кнопки, скролим данные, пожалуйста подождите!")

            process = subprocess.Popen(
                ["python", self.script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=True,
                bufsize=1,
                universal_newlines=True
            )

            self.progress_timer.start(50)  # Таймер для плавного прогресса

            while True:
                if not self._is_running:
                    process.terminate()
                    break

                output = process.stdout.readline() if process.stdout else ''
                if output == '' and process.poll() is not None:
                    break

                if output:
                    self.console_output.emit(output.strip())
                    if "Progress:" in output:
                        try:
                            progress = int(output.split("Progress:")[1].split("%")[0].strip())
                            self.current_progress = max(0, min(100, progress))
                            self.progress.emit(self.current_progress)
                        except Exception as e:
                            self.console_output.emit(f"Ошибка парсинга прогресса: {e}")

            self.progress_timer.stop()
            return_code = process.wait()

            if return_code == 0:
                self.finished.emit("Завершено успешно", True)
            else:
                error_msg = process.stderr.read() if process.stderr else "Неизвестная ошибка"
                raise subprocess.CalledProcessError(return_code, self.script_path, error_msg)

        except Exception as e:
            self.error.emit(str(e))
            self.finished.emit(f"Ошибка: {str(e)}", False)

    def smooth_progress(self):
        """Плавное увеличение значения прогресс-бара"""
        if self.last_progress < self.current_progress:
            self.last_progress += 1
            self.progress.emit(self.last_progress)

    def stop(self):
        """Остановка парсера"""
        self._is_running = False
        self.progress_timer.stop()
        self.terminate()

class NewsParserUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер новостей Алены")
        self.setWindowIcon(QIcon("SEV.ico"))
        self.setFixedSize(600, 700)

        # Инициализация переменных
        self.script_status = QLabel()
        self.loading_movie = QMovie("loading.gif")
        self.loading_movie.setScaledSize(QSize(100, 100))
        self.loading_label = QLabel()
        self.loading_label.setMovie(self.loading_movie)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_label.hide()
        self.current_worker = None
        self.ready_message = "Готов к работе"

        self.load_fonts()
        self.init_ui()
        self.set_gradient_background()

    def load_fonts(self):
        """Загрузка пользовательских шрифтов"""
        font_dir = os.path.join(os.path.dirname(__file__), "fonts")
        if os.path.exists(font_dir):
            for font_file in os.listdir(font_dir):
                if font_file.endswith(('.ttf', '.otf')):
                    QFontDatabase.addApplicationFont(os.path.join(font_dir, font_file))

    def set_gradient_background(self):
        """Установка градиентного фона окна"""
        palette = QPalette()
        gradient = QLinearGradient(0, 0, 0, self.height())
        gradient.setColorAt(0.0, QColor("#4D0B61"))
        gradient.setColorAt(1.0, QColor("#285E4B"))
        palette.setBrush(QPalette.ColorRole.Window, QBrush(gradient))
        self.setPalette(palette)

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Настройка шрифта для кнопок
        din_pro = QFont("DIN Pro Black", 12)
        din_pro.setWeight(QFont.Weight.Black)

        # Стиль для кнопок
        button_style = """
            QPushButton {
                background-color: #fff12b;
                color: black;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-family: 'DIN Pro Black';
            }
            QPushButton:hover {
                background-color: #fff12b;
            }
            QPushButton:disabled {
                background-color: #aaaaaa;
            }
        """

        # Логотип приложения
        self.logo = QLabel()
        logo_paths = ["SEV.png", os.path.join(os.path.dirname(__file__), "SEV.png")]
        for path in logo_paths:
            if os.path.exists(path):
                pixmap = QPixmap(path).scaled(150, 150,
                                            Qt.AspectRatioMode.KeepAspectRatio,
                                            Qt.TransformationMode.SmoothTransformation)
                self.logo.setPixmap(pixmap)
                break

        self.logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.logo)

        # Статус доступности скриптов
        self.script_status.setFont(din_pro)
        self.script_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.check_scripts_availability()
        main_layout.addWidget(self.script_status)

        # Контейнер для кнопок в 2 столбца
        buttons_container = QWidget()
        buttons_layout = QHBoxLayout()
        buttons_container.setLayout(buttons_layout)

        # Создаем два вертикальных столбца для кнопок
        left_column = QVBoxLayout()
        right_column = QVBoxLayout()

        # Создаем и распределяем кнопки по столбцам
        self.buttons = {}
        for i, name in enumerate(SCRIPTS):
            btn = QPushButton(name)
            btn.setFont(din_pro)
            btn.setStyleSheet(button_style)
            btn.setMinimumHeight(45)
            btn.clicked.connect(lambda checked, n=name: self.run_parser(n))

            # Распределяем кнопки по столбцам
            if i % 2 == 0:
                left_column.addWidget(btn)
            else:
                right_column.addWidget(btn)

            self.buttons[name] = btn

        # Добавляем столбцы в горизонтальный контейнер
        buttons_layout.addLayout(left_column)
        buttons_layout.addLayout(right_column)

        # Добавляем контейнер с кнопками в основной layout
        main_layout.addWidget(buttons_container)

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFont(din_pro)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #315045;
                border-radius: 5px;
                text-align: center;
                font-family: 'DIN Pro Black';
            }
            QProgressBar::chunk {
                background-color: #16f062;
                width: 10px;
            }
        """)
        main_layout.addWidget(self.progress_bar)

        # Кнопка остановки
        self.stop_button = QPushButton("Остановить")
        self.stop_button.setFont(din_pro)
        self.stop_button.setStyleSheet(button_style)
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_parser)
        main_layout.addWidget(self.stop_button)

        # Консольный вывод
        self.console_output = QPlainTextEdit()
        self.console_output.setReadOnly(True)
        self.console_output.setFont(QFont("Consolas", 10))
        self.console_output.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1a1a1a;
                color: #ffffff;
                border: 1px solid #315045;
                border-radius: 5px;
            }
        """)
        self.console_output.setMaximumHeight(200)
        main_layout.addWidget(self.console_output)

        # Статусная строка
        self.status = QStatusBar()
        self.status.setStyleSheet("""
            QStatusBar {
                background-color: #315045;
                color: white;
                font-family: 'DIN Pro Black';
                font-size: 12px;
            }
        """)
        self.setStatusBar(self.status)
        self.status.showMessage(self.ready_message)

        # Основной контейнер
        container = QWidget()
        container.setLayout(main_layout)

        # Добавляем скроллинг
        scroll = QScrollArea()
        scroll.setWidget(container)
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)

        # Анимация загрузки
        main_layout.addWidget(self.loading_label)

    def check_scripts_availability(self):
        """Проверка доступности скриптов парсеров"""
        unavailable = []
        for name, path in SCRIPTS.items():
            if not os.path.exists(path):
                unavailable.append(name)

        if unavailable:
            self.script_status.setText(f"⚠️ Не найдены скрипты: {', '.join(unavailable)}")
            self.script_status.setStyleSheet("color: #cf3c28; font-weight: bold;")
        else:
            self.script_status.setText("✓ Все скрипты доступны")
            self.script_status.setStyleSheet("color: #1eeb74; font-weight: bold;")

    def run_parser(self, parser_name):
        """Запуск выбранного парсера"""
        script = SCRIPTS[parser_name]

        if not os.path.exists(script):
            QMessageBox.critical(self, "Ошибка", f"Файл скрипта не найден:\n{script}")
            return

        if self.current_worker and self.current_worker.isRunning():
            self.current_worker.stop()

        # Подготовка UI к запуску
        self.console_output.clear()
        self.clear_status_message()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.logo.hide()
        self.loading_label.show()
        self.loading_movie.start()

        # Создание и настройка потока
        self.current_worker = ParserWorker(script)
        self.current_worker.progress.connect(self.update_progress)
        self.current_worker.finished.connect(self.on_parser_finished)
        self.current_worker.error.connect(self.on_parser_error)
        self.current_worker.console_output.connect(self.update_console_output)

        # Блокировка кнопок
        for btn in self.buttons.values():
            btn.setEnabled(False)
        self.stop_button.setEnabled(True)

        # Запуск потока
        self.current_worker.start()

    def update_progress(self, progress):
        """Обновление прогресса на UI"""
        self.progress_bar.setValue(progress)
        self.progress_bar.setFormat(f"{progress}%")

    def stop_parser(self):
        """Остановка текущего парсера"""
        if self.current_worker and self.current_worker.isRunning():
            self.current_worker.stop()
            self.ready_status_message()
            self.progress_bar.setValue(0)

            # Разблокировка кнопок
            for btn in self.buttons.values():
                btn.setEnabled(True)
            self.stop_button.setEnabled(False)

            # Сброс анимации
            self.loading_movie.stop()
            self.loading_label.hide()
            self.logo.show()

    def update_console_output(self, text):
        """Обновление вывода в консольном виджете"""
        self.console_output.appendPlainText(text)
        scrollbar = self.console_output.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def clear_status_message(self):
        """Очистка сообщения в статусной строке"""
        self.status.showMessage("")

    def ready_status_message(self):
        """Восстановление стандартного сообщения"""
        self.status.showMessage(self.ready_message)

    def on_parser_finished(self, message, success):
        """Обработчик завершения работы парсера"""
        color = "#1eeb74" if success else "#cf3c28"
        self.status.showMessage(f"{message}")
        self.status.setStyleSheet(f"color: {color};")

        # Разблокировка кнопок
        for btn in self.buttons.values():
            btn.setEnabled(True)
        self.stop_button.setEnabled(False)

        # Сброс анимации
        self.loading_movie.stop()
        self.loading_label.hide()
        self.logo.show()

        # Восстановление стандартного сообщения через 3 секунды
        QTimer.singleShot(3000, self.ready_status_message)

    def on_parser_error(self, error):
        """Обработчик ошибок парсера"""
        QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n{error[:500]}")
        self.ready_status_message()

        # Сброс анимации
        self.loading_movie.stop()
        self.loading_label.hide()
        self.logo.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = NewsParserUI()
    window.show()
    sys.exit(app.exec())