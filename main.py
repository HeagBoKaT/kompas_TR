import os
import sys
import json
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QGroupBox,
    QTreeWidget,
    QTreeWidgetItem,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QTabWidget,
    QListWidget,
    QLabel,
    QStatusBar,
    QToolBar,
    QMenuBar,
    QMenu,
    QMessageBox,
    QInputDialog,
    QScrollBar,
)

from PyQt6.QtGui import QIcon, QFont, QTextCharFormat, QTextCursor, QAction
from PyQt6.QtCore import Qt, QTimer
import pythoncom
import win32com.client
from win32com.client import Dispatch, gencache
import re
import gc


class KompasApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.status_bar = self.statusBar()  # Инициализация статусной строки
        self.status_bar.showMessage("Приложение запущено")  # Теперь работает
        self.setWindowTitle("Редактор технических требований KOMPAS-3D")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1000, 700)

        # Установка иконки приложения
        icon_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "icons", "icon.ico"
        )
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Инициализация переменных
        self.templates = {}
        self.template_search_var = ""
        self.auto_numbering_var = False

        # Загрузка шаблонов
        self.load_templates()

        # Создание пользовательского интерфейса
        self.create_ui()

        # Подключение к API Kompas
        self.module7 = None
        self.api7 = None
        self.const7 = None
        self.app7 = None
        self.connect_to_kompas()

        # Обновление информации
        self.update_active_document_info()
        self.update_documents_tree()

        # Периодическое обновление информации о документах
        self.timer = QTimer()
        self.timer.timeout.connect(self.periodic_update)
        self.timer.start(1000)

    def create_ui(self):
        """Создание пользовательского интерфейса"""
        # Создание меню
        self.create_menu()

        # Создание панели инструментов
        self.create_toolbar()

        # Создание центрального виджета
        self.create_central_widget()

        # Создание строки статуса
        self.create_status_bar()

    def create_menu(self):
        """Создание главного меню"""
        menu_bar = self.menuBar()

        # Меню "Файл"
        file_menu = menu_bar.addMenu("Файл")
        connect_action = QAction("Подключиться к KOMPAS-3D", self)
        connect_action.setShortcut("Ctrl+K")
        connect_action.triggered.connect(self.connect_to_kompas)
        file_menu.addAction(connect_action)

        check_connect_action = QAction("Проверить подключение", self)
        check_connect_action.triggered.connect(self.check_kompas_connection)
        file_menu.addAction(check_connect_action)

        file_menu.addSeparator()

        get_req_action = QAction("Получить технические требования", self)
        get_req_action.setShortcut("Ctrl+Q")
        get_req_action.triggered.connect(self.get_technical_requirements)
        file_menu.addAction(get_req_action)

        save_req_action = QAction("Сохранить технические требования", self)
        save_req_action.setShortcut("Ctrl+S")
        save_req_action.triggered.connect(self.save_technical_requirements)
        file_menu.addAction(save_req_action)

        apply_req_action = QAction("Применить технические требования", self)
        apply_req_action.setShortcut("Ctrl+E")
        apply_req_action.triggered.connect(lambda: self.apply_technical_requirements())
        file_menu.addAction(apply_req_action)

        file_menu.addSeparator()

        disconnect_action = QAction("Отключиться от KOMPAS-3D", self)
        disconnect_action.triggered.connect(self.disconnect_from_kompas)
        file_menu.addAction(disconnect_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Меню "Инструменты"
        tools_menu = menu_bar.addMenu("Инструменты")
        edit_templates_action = QAction("Редактировать файл шаблонов", self)
        edit_templates_action.triggered.connect(self.edit_templates_file)
        tools_menu.addAction(edit_templates_action)

        reload_templates_action = QAction("Обновить шаблоны", self)
        reload_templates_action.setShortcut("F5")
        reload_templates_action.triggered.connect(self.reload_templates)
        tools_menu.addAction(reload_templates_action)

        tools_menu.addSeparator()

        refresh_docs_action = QAction("Обновить список документов", self)
        refresh_docs_action.setShortcut("F6")
        refresh_docs_action.triggered.connect(self.update_documents_tree)
        tools_menu.addAction(refresh_docs_action)

        # Меню "Помощь"
        help_menu = menu_bar.addMenu("Помощь")
        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        shortcuts_action = QAction("Горячие клавиши", self)
        shortcuts_action.triggered.connect(self.show_shortcuts)
        help_menu.addAction(shortcuts_action)

    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = self.addToolBar("Toolbar")
        toolbar.setMovable(False)

        # Кнопка подключения
        connect_btn = QAction("🔌", self)
        connect_btn.setToolTip("Подключиться к KOMPAS-3D (Ctrl+K)")
        connect_btn.triggered.connect(self.connect_to_kompas)
        toolbar.addAction(connect_btn)

        # Кнопка отключения
        disconnect_btn = QAction("🚫", self)
        disconnect_btn.setToolTip("Отключиться от KOMPAS-3D")
        disconnect_btn.triggered.connect(self.disconnect_from_kompas)
        toolbar.addAction(disconnect_btn)

        # Кнопка проверки подключения
        check_connect_btn = QAction("🔍", self)
        check_connect_btn.setToolTip("Проверить подключение к KOMPAS-3D")
        check_connect_btn.triggered.connect(self.check_kompas_connection)
        toolbar.addAction(check_connect_btn)

        # Кнопка обновления списка документов
        refresh_btn = QAction("🔄", self)
        refresh_btn.setToolTip("Обновить список документов (F6)")
        refresh_btn.triggered.connect(self.update_documents_tree)
        toolbar.addAction(refresh_btn)

        toolbar.addSeparator()

        # Кнопка получения тех. требований
        get_btn = QAction("📥", self)
        get_btn.setToolTip("Получить технические требования (Ctrl+Q)")
        get_btn.triggered.connect(self.get_technical_requirements)
        toolbar.addAction(get_btn)

        # Кнопка сохранения тех. требований
        save_btn = QAction("💾", self)
        save_btn.setToolTip("Сохранить технические требования (Ctrl+S)")
        save_btn.triggered.connect(self.save_technical_requirements)
        toolbar.addAction(save_btn)

        # Кнопка применения тех. требований
        apply_btn = QAction("🔄", self)
        apply_btn.setToolTip("Применить технические требования (Ctrl+E)")
        apply_btn.triggered.connect(lambda: self.apply_technical_requirements())
        toolbar.addAction(apply_btn)

        toolbar.addSeparator()

        # Кнопка редактирования шаблонов
        edit_templates_btn = QAction("📝", self)
        edit_templates_btn.setToolTip("Редактировать файл шаблонов")
        edit_templates_btn.triggered.connect(self.edit_templates_file)
        toolbar.addAction(edit_templates_btn)

        # Кнопка обновления шаблонов
        reload_templates_btn = QAction("📋", self)
        reload_templates_btn.setToolTip("Обновить шаблоны (F5)")
        reload_templates_btn.triggered.connect(self.reload_templates)
        toolbar.addAction(reload_templates_btn)

    def create_central_widget(self):
        """Создание центрального виджета"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Информация о документе
        doc_frame = QGroupBox("Активный документ")
        doc_layout = QHBoxLayout(doc_frame)
        self.connect_status = QLabel("🔴 Нет подключения")
        self.connect_status.setStyleSheet("color: red;")
        doc_layout.addWidget(self.connect_status)
        self.active_doc_label = QLabel("Нет активного документа")
        self.active_doc_label.setWordWrap(True)
        doc_layout.addWidget(self.active_doc_label)
        doc_frame.setFixedHeight(50)  # Устанавливаем фиксированную высоту 50 пикселей
        main_layout.addWidget(doc_frame)

        # Разделение на панели
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Левая панель - дерево документов
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)

        # Правая панель - шаблоны и редактор
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)

    def create_left_panel(self):
        """Создание левой панели с деревом документов"""
        left_panel = QGroupBox("Открытые документы")
        left_layout = QVBoxLayout(left_panel)

        # Панель поиска
        search_layout = QHBoxLayout()
        search_label = QLabel("🔍")
        self.doc_search_edit = QLineEdit()
        self.doc_search_edit.setPlaceholderText("Поиск документов...")
        self.doc_search_edit.textChanged.connect(self.filter_documents_tree)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.doc_search_edit)

        refresh_btn = QPushButton("🔄")
        refresh_btn.setFixedWidth(30)
        refresh_btn.clicked.connect(self.update_documents_tree)
        refresh_btn.setToolTip("Обновить список документов (F6)")
        search_layout.addWidget(refresh_btn)
        left_layout.addLayout(search_layout)

        # Дерево документов
        self.doc_tree = QTreeWidget()
        self.doc_tree.setHeaderLabels(["Имя", "Тип", "Путь"])
        self.doc_tree.setColumnWidth(0, 150)
        self.doc_tree.setColumnWidth(1, 100)
        self.doc_tree.setColumnWidth(2, 300)
        self.doc_tree.itemDoubleClicked.connect(self.on_document_double_click)
        left_layout.addWidget(self.doc_tree)

        return left_panel

    def create_right_panel(self):
        """Создание правой панели с шаблонами и редактором"""
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        # Блок шаблонов
        templates_frame = QGroupBox("Шаблоны технических требований")
        templates_layout = QVBoxLayout(templates_frame)

        # Поиск шаблонов
        search_layout = QHBoxLayout()
        search_label = QLabel("🔍")
        self.template_search_edit = QLineEdit()
        self.template_search_edit.setPlaceholderText("Поиск шаблонов...")
        self.template_search_edit.textChanged.connect(self.filter_templates)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.template_search_edit)
        templates_layout.addLayout(search_layout)

        # Вкладки шаблонов
        self.template_tabs = QTabWidget()
        self.populate_template_tabs()
        templates_layout.addWidget(self.template_tabs)
        right_layout.addWidget(templates_frame)

        # Блок редактора
        editor_frame = QGroupBox("Текущие технические требования")
        editor_layout = QVBoxLayout(editor_frame)

        self.current_reqs_text = QTextEdit()
        self.current_reqs_text.setAcceptRichText(True)
        editor_layout.addWidget(self.current_reqs_text)

        right_layout.addWidget(editor_frame)

        return right_panel

    def create_status_bar(self):
        """Создание строки статуса"""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готово")

        self.docs_count_label = QLabel("Документов: 0")
        self.status_bar.addPermanentWidget(self.docs_count_label)

        version_label = QLabel("v1.0 (2025)")
        self.status_bar.addPermanentWidget(version_label)

    def load_templates(self):
        """Загрузка шаблонов технических требований из файла JSON"""
        try:
            user_home = os.path.expanduser("~")
            app_folder = os.path.join(user_home, "KOMPAS-TR")
            if not os.path.exists(app_folder):
                os.makedirs(app_folder)
            self.templates_file = os.path.join(app_folder, "templates.json")

            if not os.path.exists(self.templates_file):
                self.status_bar.showMessage("Файл шаблонов не найден, создаем новый")
                old_templates_file = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "templates.json"
                )
                if os.path.exists(old_templates_file):
                    with open(old_templates_file, "r", encoding="utf-8") as f_old:
                        templates_data = json.load(f_old)
                    with open(self.templates_file, "w", encoding="utf-8") as f_new:
                        json.dump(templates_data, f_new, ensure_ascii=False, indent=4)
                    self.status_bar.showMessage(
                        "Файл шаблонов перенесен в папку пользователя"
                    )
                else:
                    with open(self.templates_file, "w", encoding="utf-8") as f:
                        json.dump({"Общие": []}, f, ensure_ascii=False, indent=4)

            with open(self.templates_file, "r", encoding="utf-8") as f:
                self.templates = json.load(f)
            self.status_bar.showMessage(
                f"Загружено {sum(len(templates) for templates in self.templates.values())} шаблонов"
            )
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка загрузки шаблонов: {str(e)}")
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось загрузить шаблоны: {str(e)}"
            )
            self.templates = {"Общие": []}

    def connect_to_kompas(self):
        """Подключение к KOMPAS-3D"""
        try:
            if hasattr(self, "app7") and self.app7:
                try:
                    app_name = self.app7.ApplicationName(FullName=False)
                    self.connect_status.setText("🟢 Подключено")
                    self.connect_status.setStyleSheet("color: green;")
                    self.status_bar.showMessage(f"Уже подключено к {app_name}")
                    return True
                except Exception:
                    self.app7 = None
                    self.status_bar.showMessage(
                        "Ошибка подключения, пробуем переподключиться..."
                    )

            try:
                self.status_bar.showMessage(
                    "Попытка подключения к запущенному KOMPAS-3D..."
                )
                self.app7 = win32com.client.Dispatch("Kompas.Application.7")
                app_name = self.app7.ApplicationName(FullName=False)
                self.module7, self.api7, self.const7 = self.get_kompas_api7()
                self.connect_status.setText("🟢 Подключено")
                self.connect_status.setStyleSheet("color: green;")
                self.status_bar.showMessage(f"Подключено к запущенному {app_name}")
                self.update_documents_tree()
                return True
            except Exception:
                try:
                    self.status_bar.showMessage("Попытка запуска KOMPAS-3D...")
                    self.app7 = win32com.client.Dispatch("Kompas.Application.7")
                    self.app7.Visible = True
                    self.app7.HideMessage = True
                    self.module7, self.api7, self.const7 = self.get_kompas_api7()
                    app_name = self.app7.ApplicationName(FullName=False)
                    self.connect_status.setText("🟢 Подключено")
                    self.connect_status.setStyleSheet("color: green;")
                    self.status_bar.showMessage(f"Запущен и подключен {app_name}")
                    self.update_documents_tree()
                    return True
                except Exception as e:
                    self.connect_status.setText("🔴 Нет подключения")
                    self.connect_status.setStyleSheet("color: red;")
                    error_message = self.handle_kompas_error(e, "подключения")
                    self.status_bar.showMessage("Не удалось подключиться к KOMPAS-3D")
                    QMessageBox.critical(self, "Ошибка подключения", error_message)
                    return False
        except Exception as e:
            self.connect_status.setText("🔴 Нет подключения")
            self.connect_status.setStyleSheet("color: red;")
            error_message = self.handle_kompas_error(e, "подключения")
            self.status_bar.showMessage("Ошибка при подключении к KOMPAS-3D")
            QMessageBox.critical(self, "Ошибка подключения", error_message)
            return False

    def check_kompas_connection(self):
        """Проверка подключения к KOMPAS-3D с выводом сообщения"""
        if self.is_kompas_running():
            app_name = self.app7.ApplicationName(FullName=True)
            version = self.app7.ApplicationVersion()
            QMessageBox.information(
                self,
                "Информация о подключении",
                f"Подключено к KOMPAS-3D\n\nПриложение: {app_name}\nВерсия: {version}",
            )
            self.status_bar.showMessage(f"Подключено к {app_name} версии {version}")
            return True
        else:
            reply = QMessageBox.question(
                self,
                "Нет подключения",
                "Нет подключения к KOMPAS-3D.\n\nХотите попробовать подключиться?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                return self.connect_to_kompas()
            return False

    def get_kompas_api7(self):
        """Получение объектов API Kompas 3D версии 7"""
        module = gencache.EnsureModule(
            "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
        )
        api = module.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
            )
        )
        const = gencache.EnsureModule(
            "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
        ).constants
        return module, api, const

    def is_kompas_running(self):
        """Проверка подключения к KOMPAS-3D"""
        try:
            return hasattr(self, "app7") and self.app7 is not None
        except:
            return False

    def filter_documents_tree(self, text):
        """Фильтрация дерева документов по поисковому запросу"""
        self.update_documents_tree(text)

    def filter_templates(self, text):
        """Фильтрация шаблонов по поисковому запросу"""
        self.populate_template_tabs(text)

    def activate_selected_document(self):
        """Активация выбранного документа в дереве"""
        selected_items = self.doc_tree.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("Нет выбранного документа")
            return False

        try:
            if not hasattr(self, "app7") or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, "app7") or not self.app7:
                    return False

            doc_name = selected_items[0].text(0)
            documents = self.app7.Documents
            for i in range(documents.Count):
                doc = documents.Item(i)
                if doc.Name == doc_name:
                    doc.Active = True
                    self.update_active_document_info()
                    self.status_bar.showMessage(f"Документ {doc_name} активирован")
                    return True
            self.status_bar.showMessage(
                f"Документ {doc_name} не найден в списке открытых документов"
            )
            return False
        except Exception as e:
            error_message = self.handle_kompas_error(e, "активации документа")
            self.status_bar.showMessage("Ошибка при активации документа")
            QMessageBox.critical(self, "Ошибка", error_message)
            return False

    def show_document_info(self):
        """Показать подробную информацию о выбранном документе"""
        selected_items = self.doc_tree.selectedItems()
        if selected_items:
            item = selected_items[0]
            doc_name = item.text(0)
            doc_type = item.text(1)
            doc_path = item.text(2)
            info = f"Имя: {doc_name}\nТип: {doc_type}\nПуть: {doc_path}"
            QMessageBox.information(self, "Информация о документе", info)

    def edit_templates_file(self):
        """Открытие файла шаблонов во внешнем редакторе"""
        try:
            if not os.path.exists(self.templates_file):
                self.status_bar.showMessage("Файл шаблонов не найден, создаем новый")
                with open(self.templates_file, "w", encoding="utf-8") as f:
                    json.dump({"Общие": []}, f, ensure_ascii=False, indent=4)

            os.startfile(self.templates_file)
            self.status_bar.showMessage(
                f"Файл шаблонов открыт для редактирования: {self.templates_file}"
            )

            reply = QMessageBox.question(
                self,
                "Обновление шаблонов",
                "После завершения редактирования файла шаблонов, "
                "хотите ли вы обновить шаблоны в программе?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                QTimer.singleShot(1000, self.reload_templates)
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при открытии файла шаблонов: {str(e)}")
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось открыть файл шаблонов: {str(e)}"
            )

    def reload_templates(self):
        """Перезагрузка шаблонов из файла"""
        try:
            current_search = self.template_search_edit.text()
            self.load_templates()
            self.populate_template_tabs()
            if current_search:
                self.template_search_edit.setText(current_search)
                self.filter_templates(current_search)
            self.status_bar.showMessage("Шаблоны успешно обновлены")
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при обновлении шаблонов: {str(e)}")
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось обновить шаблоны: {str(e)}"
            )

    def show_about(self):
        """Показать информацию о программе"""
        about_text = """
        Редактор технических требований для KOMPAS-3D
        
        Программа для редактирования и вставки типовых 
        текстов в технические требования чертежей KOMPAS-3D.
        
        2025
        """
        QMessageBox.information(self, "О программе", about_text)

    def show_shortcuts(self):
        """Показать горячие клавиши"""
        shortcuts_text = """
        Горячие клавиши:
        Ctrl+K - Подключиться к KOMPAS-3D
        Ctrl+Q - Получить технические требования
        Ctrl+S - Сохранить технические требования
        Ctrl+E - Применить технические требования
        F5 - Обновить шаблоны
        F6 - Обновить список документов
        """
        QMessageBox.information(self, "Горячие клавиши", shortcuts_text)

    def populate_template_tabs(self, search_term=None):
        """Заполнение вкладок шаблонами технических требований"""
        self.template_tabs.clear()

        # Вкладка "Все"
        all_tab = QWidget()
        all_layout = QVBoxLayout(all_tab)
        all_list = QListWidget()
        all_layout.addWidget(all_list)
        self.template_tabs.addTab(all_tab, "Все")

        found_count = 0

        for category, templates in self.templates.items():
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            list_widget = QListWidget()
            tab_layout.addWidget(list_widget)
            self.template_tabs.addTab(tab, category)

            category_found_count = 0
            for template in templates:
                if (
                    search_term is None
                    or not search_term
                    or search_term.lower() in template.lower()
                    or search_term.lower() in category.lower()
                ):
                    list_widget.addItem(template)
                    all_list.addItem(f"[{category}] {template}")
                    category_found_count += 1
                    found_count += 1

            list_widget.itemDoubleClicked.connect(
                lambda item, lw=list_widget: self.insert_template(item.text())
            )

        if search_term:
            self.template_tabs.setCurrentIndex(0)
            self.status_bar.showMessage(
                f"Найдено шаблонов: {found_count} по запросу '{search_term}'"
            )
        else:
            self.status_bar.showMessage("Показаны все шаблоны")

    def update_active_document_info(self):
        """Обновление информации об активном документе"""
        try:
            if not hasattr(self, "app7") or not self.app7:
                self.connect_status.setText("🔴 Нет подключения")
                self.connect_status.setStyleSheet("color: red;")
                self.active_doc_label.setText("Нет активного документа")
                self.status_bar.showMessage("Нет подключения к KOMPAS-3D")
                return

            active_doc = self.app7.ActiveDocument
            if active_doc:
                doc_name = active_doc.Name
                doc_type = "Неизвестный тип"
                try:
                    doc2D_s = active_doc._oleobj_.QueryInterface(
                        self.module7.NamesToIIDMap["IDrawingDocument"],
                        pythoncom.IID_IDispatch,
                    )
                    doc_type = "Чертеж"
                except:
                    try:
                        doc3D_s = active_doc._oleobj_.QueryInterface(
                            self.module7.NamesToIIDMap["IDocument3D"],
                            pythoncom.IID_IDispatch,
                        )
                        doc_type = "3D-модель"
                    except:
                        try:
                            spec_s = active_doc._oleobj_.QueryInterface(
                                self.module7.NamesToIIDMap["ISpecificationDocument"],
                                pythoncom.IID_IDispatch,
                            )
                            doc_type = "Спецификация"
                        except:
                            pass
                doc_path = active_doc.Path or "Документ не сохранен"
                self.active_doc_label.setText(f"Документ: {doc_name} ({doc_type})")
                self.connect_status.setText("🟢 Подключено")
                self.connect_status.setStyleSheet("color: green;")
                self.status_bar.showMessage(
                    f"Активный документ: {doc_name} ({doc_type}) - {doc_path}"
                )
                self.select_document_in_tree(active_doc)
            else:
                self.active_doc_label.setText("Нет активного документа")
                self.status_bar.showMessage("Нет активного документа в KOMPAS-3D")
        except Exception as e:
            self.status_bar.showMessage(
                f"Ошибка при обновлении информации о документе: {str(e)}"
            )
            self.active_doc_label.setText("Ошибка обновления информации")

    def on_document_double_click(self, item, column):
        """Обработка двойного клика на документе в дереве"""
        doc_name = item.text(0)
        doc_type = item.text(1)
        if self.activate_document_by_name(doc_name):
            if doc_type == "Чертеж":
                QTimer.singleShot(500, self.get_technical_requirements)
                self.status_bar.showMessage("Загрузка технических требований...")

    def activate_document_by_name(self, doc_name):
        """Активация документа по имени"""
        try:
            if not hasattr(self, "app7") or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, "app7") or not self.app7:
                    return False
            documents = self.app7.Documents
            for i in range(documents.Count):
                doc = documents.Item(i)
                if doc.Name == doc_name:
                    doc.Active = True
                    self.update_active_document_info()
                    self.status_bar.showMessage(f"Документ {doc_name} активирован")
                    return True
            self.status_bar.showMessage(f"Документ {doc_name} не найден")
            return False
        except Exception as e:
            error_message = self.handle_kompas_error(e, "активации документа")
            self.status_bar.showMessage("Ошибка при активации документа")
            QMessageBox.critical(self, "Ошибка", error_message)
            return False

    def insert_template(self, template_text):
        """Вставка выбранного шаблона в текстовое поле"""
        if template_text:
            self.current_reqs_text.insertPlainText(template_text + "\n")
            self.status_bar.showMessage(f"Вставлен шаблон: {template_text[:30]}...")

    def get_technical_requirements(self):
        """Получение технических требований из активного документа"""
        try:
            if not hasattr(self, "module7") or not self.module7:
                self.connect_to_kompas()
                if not hasattr(self, "module7") or not self.module7:
                    return

            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.status_bar.showMessage("Нет активного документа")
                QMessageBox.warning(
                    self, "Внимание", "Нет активного документа в КОМПАС-3D"
                )
                return

            try:
                drawing_document = self.module7.IDrawingDocument(active_doc)
                tech_demand = drawing_document.TechnicalDemand

                if not tech_demand.IsCreated:
                    self.status_bar.showMessage(
                        "В документе отсутствуют технические требования!"
                    )
                    QMessageBox.warning(
                        self,
                        "Внимание",
                        "В документе отсутствуют технические требования!",
                    )
                    return

                text = tech_demand.Text
                if text.Count == 0:
                    self.status_bar.showMessage("Технические требования пусты!")
                    QMessageBox.warning(
                        self, "Внимание", "Технические требования есть, но они пусты!"
                    )
                    return

                formatted_text = self.parse_tech_req(text)
                self.current_reqs_text.setPlainText(formatted_text)
                doc_name = active_doc.Name
                self.status_bar.showMessage(
                    f"Технические требования загружены из {doc_name}"
                )
            except Exception as e:
                error_message = self.handle_kompas_error(
                    e, "получения технических требований"
                )
                self.status_bar.showMessage("Ошибка при получении тех. требований")
                QMessageBox.critical(self, "Ошибка", error_message)
        except Exception as e:
            error_message = self.handle_kompas_error(e, "работы с документом")
            self.status_bar.showMessage("Ошибка при работе с документом")
            QMessageBox.critical(self, "Ошибка", error_message)

    def save_technical_requirements(self):
        """Сохранение технических требований в активный документ"""
        self.apply_technical_requirements(save_document=True)

    def apply_technical_requirements(self, save_document=False):
        """Применение технических требований к активному документу"""
        try:
            if not hasattr(self, "module7") or not self.module7:
                self.connect_to_kompas()
                if not hasattr(self, "module7") or not self.module7:
                    return

            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.status_bar.showMessage("Нет активного документа")
                QMessageBox.warning(
                    self, "Внимание", "Нет активного документа в КОМПАС-3D"
                )
                return

            text_content = self.current_reqs_text.toPlainText().strip()

            try:
                drawing_document = self.module7.IDrawingDocument(active_doc)
                tech_demand = drawing_document.TechnicalDemand

                if not text_content:
                    if tech_demand.IsCreated:
                        text_obj = tech_demand.Text
                        while text_obj.Count > 0:
                            line = text_obj.TextLines[0]
                            line.Delete()
                        if hasattr(tech_demand, "Update"):
                            tech_demand.Update()
                        if hasattr(drawing_document, "Update"):
                            drawing_document.Update()
                        elif hasattr(active_doc, "Update"):
                            active_doc.Update()
                        self.status_bar.showMessage("Технические требования очищены")
                    else:
                        self.status_bar.showMessage(
                            "Нет технических требований для применения"
                        )
                    return

                if not tech_demand.IsCreated:
                    tech_demand.Create()

                text_obj = tech_demand.Text
                while text_obj.Count > 0:
                    line = text_obj.TextLines[0]
                    line.Delete()

                lines = text_content.split("\n")
                lines = [line.strip() for line in lines if line.strip()]
                processed_lines = []

                if self.auto_numbering_var:
                    cleaned_lines = []
                    should_number = []
                    for i, line in enumerate(lines):
                        clean_line = re.sub(r"^\d+\.\s*", "", line).lstrip()
                        cleaned_lines.append(clean_line)
                        if i > 0 and (
                            (len(clean_line) > 0 and clean_line[0].islower())
                            or clean_line.startswith("-")
                            or clean_line.startswith("–")
                        ):
                            should_number.append(False)
                        else:
                            should_number.append(True)
                    for i, (line, should_num) in enumerate(
                        zip(cleaned_lines, should_number)
                    ):
                        processed_lines.append((line, should_num))
                else:
                    for line in lines:
                        num_match = re.match(r"^(\d+)\.\s*(.*)", line)
                        if num_match:
                            req_text = num_match.group(2).strip()
                            processed_lines.append((req_text, True))
                        else:
                            indent_match = re.match(r"^\s+(.+)", line)
                            if indent_match:
                                req_text = indent_match.group(1).strip()
                                processed_lines.append((req_text, False))
                            else:
                                processed_lines.append((line, True))

                for line_text, is_numbered in processed_lines:
                    try:
                        text_line = text_obj.Add()
                        text_line.Str = line_text
                        text_line.Numbering = 1 if is_numbered else 0
                    except Exception as line_error:
                        print(
                            f"Ошибка при добавлении строки '{line_text}': {str(line_error)}"
                        )

                if hasattr(tech_demand, "Update"):
                    tech_demand.Update()
                if hasattr(drawing_document, "Update"):
                    drawing_document.Update()
                elif hasattr(active_doc, "Update"):
                    active_doc.Update()

                if save_document:
                    try:
                        active_doc.Save()
                        self.status_bar.showMessage("Документ сохранен")
                    except Exception as e:
                        error_msg = self.handle_kompas_error(e, "сохранения документа")
                        self.status_bar.showMessage(
                            "Не удалось сохранить документ автоматически"
                        )

                doc_name = active_doc.Name
                self.status_bar.showMessage(
                    f"Технические требования применены к {doc_name}"
                    + (" и сохранены" if save_document else " (без сохранения файла)")
                )
                QMessageBox.information(
                    self,
                    "Информация",
                    f"Технические требования успешно {'сохранены' if save_document else 'применены'} в {doc_name}",
                )
            except Exception as e:
                error_message = self.handle_kompas_error(
                    e, "применения технических требований"
                )
                self.status_bar.showMessage("Ошибка при применении тех. требований")
                QMessageBox.critical(self, "Ошибка", error_message)
        except Exception as e:
            error_message = self.handle_kompas_error(e, "работы с документом")
            self.status_bar.showMessage("Ошибка при работе с документом")
            QMessageBox.critical(self, "Ошибка", error_message)

    def select_document_in_tree(self, document):
        """Выбор документа в дереве документов"""
        try:
            if not document:
                return
            doc_name = document.Name
            for i in range(self.doc_tree.topLevelItemCount()):
                item = self.doc_tree.topLevelItem(i)
                if item.text(0) == doc_name:
                    self.doc_tree.setCurrentItem(item)
                    self.doc_tree.scrollToItem(item)
                    return
            self.update_documents_tree()
            for i in range(self.doc_tree.topLevelItemCount()):
                item = self.doc_tree.topLevelItem(i)
                if item.text(0) == doc_name:
                    self.doc_tree.setCurrentItem(item)
                    self.doc_tree.scrollToItem(item)
                    return
        except Exception:
            pass

    def update_documents_tree(self, search_term=None):
        """Обновление дерева документов"""
        try:
            if not hasattr(self, "app7") or not self.app7:
                self.status_bar.showMessage("Нет подключения к KOMPAS-3D")
                return

            self.doc_tree.clear()
            documents = self.app7.Documents
            doc_count = 0

            for i in range(documents.Count):
                doc = documents.Item(i)
                doc_name = doc.Name
                if search_term and search_term.lower() not in doc_name.lower():
                    continue

                doc_type = "Неизвестный тип"
                try:
                    doc._oleobj_.QueryInterface(
                        self.module7.NamesToIIDMap["IDrawingDocument"],
                        pythoncom.IID_IDispatch,
                    )
                    doc_type = "Чертеж"
                except:
                    try:
                        doc._oleobj_.QueryInterface(
                            self.module7.NamesToIIDMap["IDocument3D"],
                            pythoncom.IID_IDispatch,
                        )
                        doc_type = "3D-модель"
                    except:
                        try:
                            doc._oleobj_.QueryInterface(
                                self.module7.NamesToIIDMap["ISpecificationDocument"],
                                pythoncom.IID_IDispatch,
                            )
                            doc_type = "Спецификация"
                        except:
                            pass

                doc_path = doc.Path or "Документ не сохранен"
                item = QTreeWidgetItem(self.doc_tree)
                item.setText(0, doc_name)
                item.setText(1, doc_type)
                item.setText(2, doc_path)

                if (
                    self.app7.ActiveDocument
                    and doc.Name == self.app7.ActiveDocument.Name
                ):
                    self.doc_tree.setCurrentItem(item)
                    self.doc_tree.scrollToItem(item)

                doc_count += 1

            self.status_bar.showMessage(f"Найдено документов: {doc_count}")
            self.docs_count_label.setText(f"Документов: {doc_count}")
        except Exception as e:
            self.status_bar.showMessage(
                f"Ошибка при обновлении дерева документов: {str(e)}"
            )
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось обновить дерево документов: {str(e)}"
            )

    def periodic_update(self):
        """Периодическое обновление информации о документах"""
        try:
            if self.is_kompas_running():
                self.update_active_document_info()
            else:
                self.connect_status.setText("🔴 Нет подключения")
                self.connect_status.setStyleSheet("color: red;")
                self.active_doc_label.setText("Нет активного документа")
        except Exception:
            pass

    def format_text(self, format_type):
        """Форматирование выделенного текста"""
        cursor = self.current_reqs_text.textCursor()
        if not cursor.hasSelection():
            self.status_bar.showMessage("Нет выделенного текста для форматирования")
            return

        fmt = QTextCharFormat()
        if format_type == "bold":
            fmt.setFontWeight(
                QFont.Weight.Bold
                if not cursor.charFormat().font().bold()
                else QFont.Weight.Normal
            )
        elif format_type == "italic":
            fmt.setFontItalic(not cursor.charFormat().font().italic())
        elif format_type == "underline":
            fmt.setFontUnderline(not cursor.charFormat().font().underline())

        cursor.mergeCharFormat(fmt)
        self.current_reqs_text.setTextCursor(cursor)
        self.status_bar.showMessage(f"Применен формат '{format_type}'")

    def create_new_document(self, doc_type="drawing"):
        """Создание нового документа в KOMPAS-3D"""
        try:
            if not hasattr(self, "app7") or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, "app7") or not self.app7:
                    return

            if doc_type == "drawing":
                doc = self.app7.Document2D()
                doc.Create(False, True)
                doc_type_name = "чертеж"
            else:
                doc = self.app7.Document3D()
                doc.Create(False, True)
                doc_type_name = "3D-модель"

            doc.Active = True
            self.update_active_document_info()
            self.update_documents_tree()
            self.status_bar.showMessage(f"Создан новый документ: {doc_type_name}")
        except Exception as e:
            error_message = self.handle_kompas_error(e, "создания нового документа")
            self.status_bar.showMessage("Ошибка при создании нового документа")
            QMessageBox.critical(self, "Ошибка", error_message)

    def show_new_document_dialog(self):
        """Показать диалог выбора типа нового документа"""
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Создание нового документа")
        dialog.setLabelText("Выберите тип документа:")
        dialog.setComboBoxItems(["Чертеж", "3D-модель"])
        dialog.setFixedSize(300, 150)
        if dialog.exec():
            choice = dialog.textValue()
            if choice == "Чертеж":
                self.create_new_document("drawing")
            elif choice == "3D-модель":
                self.create_new_document("3d")

    def disconnect_from_kompas(self):
        """Отключение от KOMPAS-3D"""
        try:
            if hasattr(self, "app7") and self.app7:
                self.app7 = None
                self.module7 = None
                self.api7 = None
                self.const7 = None
                gc.collect()
                self.connect_status.setText("🔴 Нет подключения")
                self.connect_status.setStyleSheet("color: red;")
                self.status_bar.showMessage("Отключено от KOMPAS-3D")
                self.doc_tree.clear()
                return True
            else:
                self.status_bar.showMessage("Нет активного подключения к KOMPAS-3D")
                return False
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при отключении от KOMPAS-3D: {str(e)}")
            return False

    def closeEvent(self, event):
        """Обработчик закрытия приложения"""
        try:
            if hasattr(self, "app7") and self.app7:
                self.disconnect_from_kompas()
            pythoncom.CoUninitialize()
            event.accept()
        except Exception as e:
            print(f"Ошибка при закрытии приложения: {str(e)}")
            event.accept()

    def handle_kompas_error(self, e, operation="операции"):
        """Обработка ошибок при работе с KOMPAS-3D"""
        error_msg = str(e)
        error_code = None
        if "0x" in error_msg:
            try:
                match = re.search(r"0x[0-9A-Fa-f]+", error_msg)
                if match:
                    error_code = match.group(0)
            except:
                pass

        if error_code:
            message = f"Ошибка при выполнении {operation} в KOMPAS-3D.\n\nКод ошибки: {error_code}"
            if error_code == "0x80004005":
                message += (
                    "\n\nНеуказанная ошибка. Возможно, проблема с доступом к объекту."
                )
            elif error_code == "0x80020009":
                message += "\n\nИсключение в KOMPAS-3D. Проверьте состояние документа."
            elif error_code == "0x8002000A":
                message += "\n\nНеверный индекс или параметр."
            elif error_code == "0x80020006":
                message += "\n\nНеизвестное имя или метод."
            message += "\n\nРекомендации:\n1. Убедитесь, что KOMPAS-3D запущен и работает корректно.\n2. Проверьте, что у вас есть права на редактирование документа.\n3. Попробуйте переподключиться к KOMPAS-3D."
        else:
            message = f"Ошибка при выполнении {operation} в KOMPAS-3D.\n\n{error_msg}"
        self.status_bar.showMessage(f"Ошибка: {error_msg}")
        return message

    def apply_list_formatting(self, tech_req):
        """Применение форматирования списком к техническим требованиям"""
        try:
            if hasattr(tech_req, "FormatAsList"):
                tech_req.FormatAsList()
                if hasattr(tech_req, "ListParams"):
                    tech_req.ListParams = True
                if hasattr(tech_req, "ListType"):
                    tech_req.ListType = 0
                self.apply_numbering(tech_req)
                self.status_bar.showMessage(
                    "Применено форматирование списком и нумерация (метод 1)"
                )
                return True
            elif hasattr(tech_req, "Text") and hasattr(tech_req.Text, "FormatAsList"):
                tech_req.Text.FormatAsList()
                if hasattr(tech_req.Text, "ListParams"):
                    tech_req.Text.ListParams = True
                if hasattr(tech_req.Text, "ListType"):
                    tech_req.Text.ListType = 0
                self.apply_numbering(tech_req)
                self.status_bar.showMessage(
                    "Применено форматирование списком и нумерация (метод 2)"
                )
                return True
            else:
                try:
                    text_obj = tech_req.Text
                    if hasattr(text_obj, "ListParams"):
                        text_obj.ListParams = True
                    if hasattr(text_obj, "ListType"):
                        text_obj.ListType = 0
                    for i in range(text_obj.Count):
                        line = text_obj.TextLines[i]
                        if hasattr(line, "ListStyle"):
                            line.ListStyle = True
                        if hasattr(line, "ListParams"):
                            line.ListParams = True
                        if hasattr(line, "ListType"):
                            line.ListType = 0
                    self.apply_numbering(tech_req)
                    self.status_bar.showMessage(
                        "Применено форматирование списком и нумерация (метод 3)"
                    )
                    return True
                except Exception as e:
                    self.status_bar.showMessage(
                        f"Не удалось применить форматирование списком: {str(e)}"
                    )
                    return False
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при форматировании списком: {str(e)}")
            return False

    def apply_numbering(self, tech_req):
        """Применение нумерации к техническим требованиям средствами API KOMPAS"""
        try:
            if hasattr(tech_req, "SetNumbering"):
                tech_req.SetNumbering()
                if hasattr(tech_req, "ListParams"):
                    tech_req.ListParams = True
                if hasattr(tech_req, "ListType"):
                    tech_req.ListType = 0
                if hasattr(tech_req, "AutoNumbering"):
                    tech_req.AutoNumbering = True
                self.status_bar.showMessage("Применена нумерация (метод 1)")
                return True
            elif hasattr(tech_req, "Text") and hasattr(tech_req.Text, "SetNumbering"):
                tech_req.Text.SetNumbering()
                if hasattr(tech_req.Text, "ListParams"):
                    tech_req.Text.ListParams = True
                if hasattr(tech_req.Text, "ListType"):
                    tech_req.Text.ListType = 0
                if hasattr(tech_req.Text, "AutoNumbering"):
                    tech_req.Text.AutoNumbering = True
                self.status_bar.showMessage("Применена нумерация (метод 2)")
                return True
            elif hasattr(tech_req, "Text") and hasattr(tech_req.Text, "NumberingStyle"):
                tech_req.Text.NumberingStyle = True
                if hasattr(tech_req.Text, "ListParams"):
                    tech_req.Text.ListParams = True
                if hasattr(tech_req.Text, "ListType"):
                    tech_req.Text.ListType = 0
                if hasattr(tech_req.Text, "AutoNumbering"):
                    tech_req.Text.AutoNumbering = True
                self.status_bar.showMessage("Применена нумерация (метод 3)")
                return True
            else:
                try:
                    text_obj = tech_req.Text
                    if hasattr(text_obj, "ListParams"):
                        text_obj.ListParams = True
                    if hasattr(text_obj, "ListType"):
                        text_obj.ListType = 0
                    if hasattr(text_obj, "AutoNumbering"):
                        text_obj.AutoNumbering = True
                    for i in range(text_obj.Count):
                        line = text_obj.TextLines[i]
                        if hasattr(line, "NumberingStyle"):
                            line.NumberingStyle = True
                        elif hasattr(line, "Numbering"):
                            line.Numbering = True
                        if hasattr(line, "ListParams"):
                            line.ListParams = True
                        if hasattr(line, "ListType"):
                            line.ListType = 0
                    if hasattr(text_obj, "ListParams"):
                        text_obj.ListParams = True
                    self.status_bar.showMessage("Применена нумерация (метод 4)")
                    return True
                except Exception as e:
                    self.status_bar.showMessage(
                        f"Не удалось применить нумерацию: {str(e)}"
                    )
                    return False
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при применении нумерации: {str(e)}")
            return False

    def apply_auto_numbering(self):
        """Применение автоматической нумерации к техническим требованиям"""
        try:
            text_content = self.current_reqs_text.toPlainText().strip()
            if not text_content:
                return

            lines = text_content.split("\n")
            lines = [line.strip() for line in lines if line.strip()]
            cleaned_lines = []
            should_number = []

            for i, line in enumerate(lines):
                clean_line = re.sub(r"^\d+\.\s*", "", line)
                cleaned_lines.append(clean_line)
                if i > 0 and (
                    (len(clean_line) > 0 and clean_line[0].islower())
                    or clean_line.startswith("-")
                    or clean_line.startswith("–")
                ):
                    should_number.append(False)
                else:
                    should_number.append(True)

            result_lines = []
            number_counter = 1
            for i, (line, should_num) in enumerate(zip(cleaned_lines, should_number)):
                if should_num:
                    result_lines.append(f"{number_counter}. {line}")
                    number_counter += 1
                else:
                    result_lines.append(f"    {line}")

            self.current_reqs_text.setPlainText("\n".join(result_lines))
            self.status_bar.showMessage("Автонумерация применена")
        except Exception as e:
            self.status_bar.showMessage(
                f"Ошибка при применении автонумерации: {str(e)}"
            )
            print(f"Error applying auto numbering: {str(e)}")

    def remove_auto_numbering(self):
        """Удаление автоматической нумерации из технических требований"""
        try:
            text_content = self.current_reqs_text.toPlainText().strip()
            if not text_content:
                return

            lines = text_content.split("\n")
            result_lines = []
            for line in lines:
                clean_line = re.sub(r"^\d+\.\s*", "", line)
                result_lines.append(clean_line)

            self.current_reqs_text.setPlainText("\n".join(result_lines))
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка при удалении автонумерации: {str(e)}")
            print(f"Error removing auto numbering: {str(e)}")

    def toggle_auto_numbering(self, is_enabled=None):
        """Переключение режима автонумерации"""
        try:
            if is_enabled is None:
                is_enabled = self.auto_numbering_var
            else:
                self.auto_numbering_var = is_enabled

            if is_enabled:
                self.apply_auto_numbering()
                # Note: PyQt6 does not support direct key binding like Tkinter; consider adding toolbar/menu actions
            else:
                self.remove_auto_numbering()
        except Exception as e:
            self.status_bar.showMessage(
                f"Ошибка при переключении автонумерации: {str(e)}"
            )
            print(f"Error toggling auto numbering: {str(e)}")

    def parse_tech_req(self, text_lines):
        """Парсинг технических требований из объекта TextLines"""
        formatted_text = ""
        count = 0
        current_req = ""
        current_req_num = 0

        i = 0
        while i < text_lines.Count:
            line = text_lines.TextLines[i]
            line_text = line.Str.strip()
            if not line_text:
                i += 1
                continue

            if line.Numbering == 1:
                if current_req:
                    formatted_text += f"{current_req_num}. {current_req}\n"
                count += 1
                current_req_num = count
                current_req = line_text
            else:
                if current_req:
                    if (
                        not current_req.endswith(" ")
                        and not current_req.endswith("-")
                        and not line_text.startswith("-")
                    ):
                        current_req += " "
                    current_req += line_text
                else:
                    count += 1
                    current_req_num = count
                    current_req = line_text

            if i == text_lines.Count - 1 and current_req:
                formatted_text += f"{current_req_num}. {current_req}\n"
            i += 1

        return formatted_text

    def clean_tech_req_line(self, line):
        """Очистка строки технических требований от нумерации и форматирования"""
        line = re.sub(r"^\s*\d+\.\s*", "", line)
        line = re.sub(r"^\s*[•\-–—]\s*", "", line)
        return line.strip()


if __name__ == "__main__":
    pythoncom.CoInitialize()
    try:
        app = QApplication(sys.argv)
        window = KompasApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        QMessageBox.critical(None, "Ошибка", f"Критическая ошибка приложения: {str(e)}")
    finally:
        pythoncom.CoUninitialize()
