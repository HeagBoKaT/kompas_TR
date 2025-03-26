import os
import sys
import json
import logging
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication,
    QHeaderView,
    QListWidgetItem,
    QMainWindow,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
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
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QLineEdit,
    QComboBox,
    QTextEdit,
    QLabel,
    QMessageBox,
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
        # Инициализация пути к файлу настроек
        user_home = os.path.expanduser("~")
        app_folder = os.path.join(user_home, "KOMPAS-TR")
        if not os.path.exists(app_folder):
            os.makedirs(app_folder)
        self.settings_file = os.path.join(app_folder, "settings.json")

        # Загружаем настройки темы (по умолчанию светлая)
        self.dark_mode = self.load_theme_setting()

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Приложение запущено")
        self.setWindowTitle("Редактор технических требований KOMPAS-3D")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1000, 700)

        # Установка иконки приложения
        icon_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "icons", "icon.ico"
        )
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.templates = {}
        self.template_search_var = ""
        self.auto_numbering_var = False

        self.load_templates()
        ThemeManager.apply_theme(self, self.dark_mode)  # Применяем загруженную тему
        self.create_ui()

        self.module7 = None
        self.api7 = None
        self.const7 = None
        self.app7 = None
        self.connect_to_kompas()

        self.update_active_document_info()
        self.update_documents_tree()

        self.timer = QTimer()
        self.timer.timeout.connect(self.periodic_update)
        self.timer.start(1000)

    def apply_theme(self):
        ThemeManager.apply_theme(self, self.dark_mode)

    def load_theme_setting(self):
        """Загрузка настройки темы из файла"""
        try:
            if not os.path.exists(self.settings_file):
                # Если файла нет, создаем с темой по умолчанию (светлая)
                default_settings = {"dark_mode": False}
                with open(self.settings_file, "w", encoding="utf-8") as f:
                    json.dump(default_settings, f, ensure_ascii=False, indent=4)
                return False

            with open(self.settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
                return settings.get("dark_mode", False)  # По умолчанию светлая
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка загрузки настроек темы: {str(e)}")
            # В случае ошибки возвращаем светлую тему
            return False

    def save_theme_setting(self):
        """Сохранение настройки темы в файл"""
        try:
            settings = {"dark_mode": self.dark_mode}
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            self.status_bar.showMessage("Настройки темы сохранены")
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка сохранения настроек темы: {str(e)}")

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

        # Добавляем действие для сохранения в PDF
        save_pdf_action = QAction("Сохранить в PDF", self)
        save_pdf_action.setShortcut("Ctrl+Shift+S")
        save_pdf_action.triggered.connect(self.save_to_pdf)
        file_menu.addAction(save_pdf_action)

        file_menu.addSeparator()

        disconnect_action = QAction("Отключиться от KOMPAS-3D", self)
        disconnect_action.triggered.connect(self.disconnect_from_kompas)
        file_menu.addAction(disconnect_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Меню "Инструменты" (без изменений)
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

        theme_action = QAction("Переключить тему", self)
        theme_action.setShortcut("Ctrl+T")
        theme_action.triggered.connect(self.toggle_theme)
        tools_menu.addAction(theme_action)

        # Меню "Помощь" (без изменений)
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

        # Кнопка сохранения в PDF
        save_pdf_btn = QAction("📄", self)
        save_pdf_btn.setToolTip("Сохранить в PDF (Ctrl+Shift+S)")
        save_pdf_btn.triggered.connect(self.save_to_pdf)
        toolbar.addAction(save_pdf_btn)

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
        # Убираем фиксированную высоту или увеличиваем
        doc_frame.setMinimumHeight(
            40
        )  # Устанавливаем минимальную высоту вместо фиксированной
        doc_frame.setMaximumHeight(
            70
        )  # Устанавливаем минимальную высоту вместо фиксированной
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

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 4)

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

        # Добавляем визуальные улучшения
        self.status_bar.setStyleSheet(
            """
            QStatusBar::item {
                border: none;
            }
            QLabel {
                padding: 4px 8px;
            }
        """
        )

        self.docs_count_label = QLabel("Документов: 0")
        self.status_bar.addPermanentWidget(self.docs_count_label)

        version_label = QLabel("v1.0.4 (2025)")
        self.status_bar.addPermanentWidget(version_label)

    def load_templates(self):
        try:
            user_home = os.path.expanduser("~")
            app_folder = os.path.join(user_home, "KOMPAS-TR")
            if not os.path.exists(app_folder):
                os.makedirs(app_folder)
            self.templates_file = os.path.join(app_folder, "templates.json")

            if not os.path.exists(self.templates_file):
                self.status_bar.showMessage("Файл шаблонов не найден, создаем новый")
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
        """Открытие редактора шаблонов"""
        dialog = TemplateEditorDialog(self, self.templates_file)
        dialog.exec()

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
        Ctrl+Shift+S - Сохранить в PDF
        F5 - Обновить шаблоны
        F6 - Обновить список документов
        """
        QMessageBox.information(self, "Горячие клавиши", shortcuts_text)

    def populate_template_tabs(self, search_term=None):
        self.template_tabs.clear()

        # Вкладка "Все"
        all_tab = QWidget()
        all_layout = QVBoxLayout(all_tab)
        all_list = QListWidget()
        all_layout.addWidget(all_list)
        self.template_tabs.addTab(all_tab, "Все")
        all_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        all_list.customContextMenuRequested.connect(
            lambda pos: self.show_template_context_menu(pos, all_list)
        )
        all_list.itemDoubleClicked.connect(self.insert_template)

        found_count = 0

        for category, templates in self.templates.items():
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            list_widget = QListWidget()
            tab_layout.addWidget(list_widget)
            self.template_tabs.addTab(tab, category)
            list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            list_widget.customContextMenuRequested.connect(
                lambda pos, lw=list_widget: self.show_template_context_menu(pos, lw)
            )
            list_widget.itemDoubleClicked.connect(self.insert_template)

            for template in templates:
                if isinstance(template, dict):
                    text = template.get("text", "")
                    variants = template.get("variants", [])
                    if (
                        search_term is None
                        or not search_term
                        or search_term.lower() in text.lower()
                        or any(
                            search_term.lower()
                            in (
                                variant.get("text", "")
                                if isinstance(variant, dict)
                                else variant
                            ).lower()
                            for variant in variants
                        )
                    ):
                        # Для категорийных вкладок
                        item = QListWidgetItem(text)
                        item.setData(Qt.ItemDataRole.UserRole, template)
                        list_widget.addItem(item)

                        # Для вкладки "Все"
                        all_item = QListWidgetItem(f"[{category}] {text}")
                        all_item.setData(Qt.ItemDataRole.UserRole, template)
                        all_list.addItem(all_item)

                        found_count += 1
                else:
                    # Обратная совместимость со старым форматом
                    if (
                        search_term is None
                        or not search_term
                        or search_term.lower() in template.lower()
                    ):
                        item = QListWidgetItem(template)
                        item.setData(
                            Qt.ItemDataRole.UserRole, {"text": template, "variants": []}
                        )
                        list_widget.addItem(item)
                        all_item = QListWidgetItem(f"[{category}] {template}")
                        all_item.setData(
                            Qt.ItemDataRole.UserRole, {"text": template, "variants": []}
                        )
                        all_list.addItem(all_item)
                        found_count += 1

        if search_term:
            self.template_tabs.setCurrentIndex(0)
            self.status_bar.showMessage(
                f"Найдено шаблонов: {found_count} по запросу '{search_term}'"
            )
        else:
            self.status_bar.showMessage("Показаны все шаблоны")

    def show_template_context_menu(self, pos, list_widget):
        item = list_widget.itemAt(pos)
        if not item:
            return

        template = item.data(Qt.ItemDataRole.UserRole)
        if (
            not isinstance(template, dict)
            or "variants" not in template
            or not template["variants"]
        ):
            return

        menu = QMenu(self)
        style = self.style()

        for variant in template["variants"]:
            if isinstance(variant, dict):
                variant_text = variant.get("text", "")
                custom_input = variant.get("custom_input", False)
            else:
                variant_text = variant
                custom_input = False

            action = QAction(variant_text, self)
            if custom_input:
                icon = style.standardIcon(
                    QStyle.StandardPixmap.SP_FileDialogDetailedView
                )
                action.setIcon(icon)

            if custom_input:
                action.triggered.connect(
                    lambda checked, t=template[
                        "text"
                    ], v=variant_text: self.insert_custom_variant(t, v)
                )
            else:
                action.triggered.connect(
                    lambda checked, t=template[
                        "text"
                    ], v=variant_text: self.insert_template_variant(t, v)
                )
            menu.addAction(action)

        menu.exec(list_widget.mapToGlobal(pos))

    def insert_custom_variant(self, base_text, variant_text):
        custom_value, ok = QInputDialog.getText(
            self, "Ввод значения", f"Введите значение для {variant_text}:"
        )
        if ok and custom_value:
            # Проверяем, есть ли в variant_text маркер {}
            if "{}" in variant_text:
                # Вставляем значение в место, указанное маркером
                full_text = f"{base_text} {variant_text.format(custom_value)}"
            else:
                # Запасной вариант: старый порядок
                full_text = f"{base_text} {custom_value} {variant_text}"
            self.current_reqs_text.insertPlainText(full_text + "\n")
            self.status_bar.showMessage(f"Вставлен шаблон: {full_text[:30]}...")

    def insert_template_variant(self, base_text, variant_text):
        full_text = f"{base_text}{variant_text}"
        self.current_reqs_text.insertPlainText(full_text + "\n")
        self.status_bar.showMessage(f"Вставлен шаблон: {full_text[:30]}...")

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

    def insert_template(self, item):
        template = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(template, dict):
            text = template.get("text", "")
            variants = template.get("variants", [])
            if variants:
                # Вставляем первую вариацию по умолчанию при двойном клике
                self.insert_template_variant(text, variants[0])
            else:
                self.current_reqs_text.insertPlainText(text + "\n")
                self.status_bar.showMessage(f"Вставлен шаблон: {text[:30]}...")
        else:
            # Обратная совместимость со старым форматом
            self.current_reqs_text.insertPlainText(template + "\n")
            self.status_bar.showMessage(f"Вставлен шаблон: {template[:30]}...")

    def get_technical_requirements(self):
        """Получение технических требований из активного документа"""
        try:
            # Проверка подключения к KOMPAS-3D
            if not hasattr(self, "module7") or not self.module7:
                self.connect_to_kompas()
                if not hasattr(self, "module7") or not self.module7:
                    return

            # Проверка наличия активного документа
            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.status_bar.showMessage("Нет активного документа")
                QMessageBox.warning(
                    self, "Внимание", "Нет активного документа в КОМПАС-3D"
                )
                return

            try:
                # Получение интерфейса чертежа и технических требований
                drawing_document = self.module7.IDrawingDocument(active_doc)
                tech_demand = drawing_document.TechnicalDemand

                # Проверка, созданы ли технические требования
                if not tech_demand.IsCreated:
                    # Создание новых пустых технических требований
                    tt = tech_demand.Text
                    stroka = tt.Add().Add()
                    stroka.Str = "  "
                    # tech_demand.Create()
                    self.status_bar.showMessage(
                        "Созданы новые пустые технические требования"
                    )
                    self.current_reqs_text.setPlainText("")
                    return

                # Получение текста технических требований
                text = tech_demand.Text
                if text.Count == 0:
                    self.status_bar.showMessage("Технические требования пусты")
                    self.current_reqs_text.setPlainText("")
                    return

                # Парсинг и отображение существующих требований
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
                print(error_message)
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
                    tt = tech_demand.Text
                    stroka = tt.Add().Add()
                    stroka.Str = "  "

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
                # Автоматическое обновление технических требований после применения
                self.get_technical_requirements()
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

    def toggle_theme(self):
        """Переключение темы с сохранением"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        # Сохраняем настройку после переключения
        self.save_theme_setting()
        # Обновляем тему для всех открытых дочерних окон
        for child in self.findChildren(TemplateEditorDialog):
            child.dark_mode = self.dark_mode
            child.apply_theme()
        self.status_bar.showMessage(
            f"Тема изменена на {'темную' if self.dark_mode else 'светлую'}"
        )

    def save_to_pdf(self):
        """Сохранение активного чертежа в PDF с пошаговым логированием"""
        try:
            # Проверка подключения к KOMPAS-3D
            if not hasattr(self, "app7") or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, "app7") or not self.app7:
                    self.status_bar.showMessage("Не удалось подключиться к KOMPAS-3D")
                    QMessageBox.critical(
                        self, "Ошибка", "Не удалось подключиться к KOMPAS-3D"
                    )
                    return

            # Проверка активного документа
            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.status_bar.showMessage("Нет активного документа")
                QMessageBox.warning(
                    self, "Ошибка", "Нет активного документа в KOMPAS-3D"
                )
                return

            doc_name = active_doc.Name

            # Проверка типа документа (должен быть чертеж)
            doc_type = active_doc.DocumentType
            if doc_type != 1:  # 1 - это тип чертежа
                self.status_bar.showMessage("Активный документ не является чертежом")
                QMessageBox.warning(
                    self, "Ошибка", "Активный документ должен быть чертежом"
                )
                return
            # Получение пути к файлу
            doc_path = active_doc.PathName
            if not doc_path:
                self.status_bar.showMessage("Документ не сохранен")
                QMessageBox.warning(self, "Ошибка", "Сначала сохраните документ")
                return

            # Формирование пути для PDF
            doc_dir = os.path.dirname(doc_path)
            doc_name_without_ext = os.path.splitext(os.path.basename(doc_path))[0]
            pdf_folder = os.path.join(doc_dir, "Чертежи в pdf")

            pdf_path = os.path.join(pdf_folder, f"{doc_name_without_ext}.pdf")
            # Получение 2D интерфейса документа
            try:
                doc_2d = win32com.client.Dispatch(active_doc, "ksDocument2D")
            except Exception as e:
                self.status_bar.showMessage("Ошибка при получении интерфейса документа")
                QMessageBox.critical(
                    self, "Ошибка", f"Не удалось получить 2D интерфейс: {str(e)}"
                )
                return

            # Сохранение в PDF
            try:
                result = doc_2d.SaveAs(pdf_path)
                if result:
                    self.status_bar.showMessage(f"Чертеж сохранен в PDF: {pdf_path}")
                    QMessageBox.information(
                        self, "Успех", f"Чертеж сохранен в PDF:\n{pdf_path}"
                    )
            except Exception as e:
                self.status_bar.showMessage("Ошибка при сохранении в PDF")
                QMessageBox.critical(
                    self, "Ошибка", f"Ошибка сохранения в PDF: {str(e)}"
                )
                return

        except Exception as e:
            error_message = self.handle_kompas_error(e, "сохранения в PDF")
            self.status_bar.showMessage("Критическая ошибка при сохранении в PDF")
            QMessageBox.critical(self, "Ошибка", error_message)


class TemplateEditorDialog(QDialog):
    def __init__(self, parent, templates_file):
        super().__init__(parent)
        self.setWindowTitle("Редактор шаблонов")
        self.setGeometry(200, 200, 1200, 700)
        self.setMinimumSize(800, 600)
        self.templates_file = templates_file
        self.templates = parent.templates.copy()
        self.selected_template = None
        self.dark_mode = parent.dark_mode  # Синхронизация с родительской темой
        ThemeManager.apply_theme(self, self.dark_mode)  # Применяем тему
        self.init_ui()

    def apply_theme(self):
        """Применение темы для диалога"""
        ThemeManager.apply_theme(self, self.dark_mode)

    def init_ui(self):
        layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая часть: дерево шаблонов
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_label = QLabel("<b>Шаблоны</b>")
        left_label.setStyleSheet("padding-bottom: 5px;")
        left_layout.addWidget(left_label)
        self.template_tree = QTreeWidget()
        self.template_tree.setHeaderLabels(["Категория", "Текст"])
        self.template_tree.setColumnWidth(0, 200)
        self.template_tree.itemClicked.connect(self.load_template_to_editor)
        left_layout.addWidget(self.template_tree)
        splitter.addWidget(left_widget)

        # Правая часть: редактор
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_label = QLabel("<b>Редактирование шаблона</b>")
        right_label.setStyleSheet("padding-bottom: 5px;")
        right_layout.addWidget(right_label)

        # Категория
        category_layout = QHBoxLayout()
        category_label = QLabel("Категория:")
        category_label.setFixedWidth(80)
        category_layout.addWidget(category_label)
        self.category_combo = QComboBox()
        self.category_combo.setEditable(True)
        self.category_combo.addItems(list(self.templates.keys()))
        category_layout.addWidget(self.category_combo)
        right_layout.addLayout(category_layout)

        # Текст шаблона
        template_label = QLabel("Текст шаблона:")
        template_label.setStyleSheet("padding-top: 5px;")
        right_layout.addWidget(template_label)
        self.template_text = QLineEdit()
        right_layout.addWidget(self.template_text)

        # Варианты
        variants_group = QGroupBox("Варианты")
        variants_layout = QVBoxLayout(variants_group)
        self.variants_table = QTableWidget()
        self.variants_table.setColumnCount(2)
        self.variants_table.setHorizontalHeaderLabels(
            ["Текст", "Пользовательский ввод"]
        )
        self.variants_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.variants_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.variants_table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked)
        self.variants_table.itemDoubleClicked.connect(self.toggle_custom_input)
        variants_layout.addWidget(self.variants_table)

        # Панель управления вариантами
        variant_controls = QHBoxLayout()
        self.variant_text = QLineEdit()
        self.variant_text.setPlaceholderText("Введите вариант")
        variant_controls.addWidget(self.variant_text)
        self.custom_input_check = QPushButton("Пользовательский ввод")
        self.custom_input_check.setCheckable(True)
        self.custom_input_check.setStyleSheet(
            """
            QPushButton:checked {
                background-color: #409EFF;
                color: white;
                border-color: #409EFF;
            }
        """
        )
        variant_controls.addWidget(self.custom_input_check)

        add_variant_btn = QPushButton("Добавить")
        add_variant_btn.clicked.connect(self.add_variant)
        variant_controls.addWidget(add_variant_btn)
        edit_variant_btn = QPushButton("Изменить")
        edit_variant_btn.clicked.connect(self.edit_variant)
        variant_controls.addWidget(edit_variant_btn)
        delete_variant_btn = QPushButton("Удалить")
        delete_variant_btn.clicked.connect(self.delete_variant)
        variant_controls.addWidget(delete_variant_btn)
        variants_layout.addLayout(variant_controls)
        right_layout.addWidget(variants_group)

        # Предпросмотр
        preview_group = QGroupBox("Предпросмотр")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        right_layout.addWidget(preview_group)

        # Кнопки управления с учетом темы
        buttons_layout = QHBoxLayout()
        self.add_button = QPushButton("Добавить шаблон")
        self.add_button.clicked.connect(self.add_template)
        buttons_layout.addWidget(self.add_button)
        self.edit_button = QPushButton("Сохранить изменения")
        self.edit_button.clicked.connect(self.edit_template)
        buttons_layout.addWidget(self.edit_button)
        self.delete_button = QPushButton("Удалить шаблон")
        self.delete_button.clicked.connect(self.delete_template)
        buttons_layout.addWidget(self.delete_button)
        self.save_button = QPushButton("Сохранить и закрыть")
        self.save_button.setObjectName("saveButton")
        self.save_button.clicked.connect(self.save_and_close)
        buttons_layout.addWidget(self.save_button)
        right_layout.addLayout(buttons_layout)

        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(2, 3)
        layout.addWidget(splitter)

        self.populate_tree()

    def populate_tree(self):
        """Заполнение дерева шаблонов"""
        self.template_tree.clear()
        for category, templates in self.templates.items():
            for template in templates:
                if isinstance(template, dict):
                    text = template.get("text", "")
                    item = QTreeWidgetItem(self.template_tree)
                    item.setText(0, category)
                    item.setText(1, text)
                    item.setData(0, Qt.ItemDataRole.UserRole, (category, template))
                else:
                    item = QTreeWidgetItem(self.template_tree)
                    item.setText(0, category)
                    item.setText(1, template)
                    item.setData(
                        0,
                        Qt.ItemDataRole.UserRole,
                        (category, {"text": template, "variants": []}),
                    )

    def load_template_to_editor(self, item):
        category, template = item.data(0, Qt.ItemDataRole.UserRole)
        self.selected_template = (category, template)
        self.category_combo.setCurrentText(category)
        self.template_text.setText(template.get("text", ""))
        self.variants_table.setRowCount(0)
        variants = template.get("variants", [])
        for variant in variants:
            row = self.variants_table.rowCount()
            self.variants_table.insertRow(row)
            text = variant.get("text", "") if isinstance(variant, dict) else variant
            custom = (
                variant.get("custom_input", False)
                if isinstance(variant, dict)
                else False
            )
            text_item = QTableWidgetItem(text)
            text_item.setFlags(
                text_item.flags() | Qt.ItemFlag.ItemIsEditable
            )  # Редактируемая
            self.variants_table.setItem(row, 0, text_item)
            custom_item = QTableWidgetItem("Да" if custom else "Нет")
            # Оставляем не редактируемой стандартным способом, переключение через обработчик
            custom_item.setFlags(custom_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.variants_table.setItem(row, 1, custom_item)

    def load_variant_details(self, item):
        """Загрузка деталей варианта в поля редактирования"""
        variant = item.data(Qt.ItemDataRole.UserRole)
        self.variant_text.setText(variant.get("text", ""))
        self.custom_input_check.setChecked(variant.get("custom_input", False))

    def add_variant(self):
        """Добавление нового варианта"""
        text = self.variant_text.text().strip()
        custom_input = self.custom_input_check.isChecked()
        if not text:
            QMessageBox.warning(self, "Ошибка", "Введите текст варианта")
            return
        row = self.variants_table.rowCount()
        self.variants_table.insertRow(row)
        text_item = QTableWidgetItem(text)
        text_item.setFlags(
            text_item.flags() | Qt.ItemFlag.ItemIsEditable
        )  # Редактируемая
        self.variants_table.setItem(row, 0, text_item)
        custom_item = QTableWidgetItem("Да" if custom_input else "Нет")
        custom_item.setFlags(
            custom_item.flags() & ~Qt.ItemFlag.ItemIsEditable
        )  # Не редактируемая
        self.variants_table.setItem(row, 1, custom_item)
        self.variant_text.clear()

    def edit_variant(self):
        """Редактирование выбранного варианта через кнопку 'Изменить'"""
        selected_row = self.variants_table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Ошибка", "Выберите вариант для изменения")
            return

        current_text = self.variants_table.item(selected_row, 0).text()
        current_custom = self.variants_table.item(selected_row, 1).text() == "Да"

        new_text, ok = QInputDialog.getText(
            self,
            "Изменить вариант",
            "Введите новый текст:",
            QLineEdit.EchoMode.Normal,
            current_text,
        )
        if ok and new_text:
            text_item = QTableWidgetItem(new_text)
            text_item.setFlags(text_item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.variants_table.setItem(selected_row, 0, text_item)
            custom_item = QTableWidgetItem("Да" if current_custom else "Нет")
            custom_item.setFlags(custom_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.variants_table.setItem(selected_row, 1, custom_item)

    def delete_variant(self):
        """Удаление выбранного варианта"""
        selected = self.variants_table.currentRow()
        if selected == -1:
            QMessageBox.warning(self, "Ошибка", "Выберите вариант для удаления")
            return
        self.variants_table.removeRow(selected)

    def add_template(self):
        """Добавление нового шаблона"""
        category = self.category_combo.currentText().strip()
        text = self.template_text.text().strip()
        if not category or not text:
            QMessageBox.warning(self, "Ошибка", "Укажите категорию и текст шаблона")
            return

        # Собираем варианты из таблицы
        variants = []
        for row in range(self.variants_table.rowCount()):
            text = self.variants_table.item(row, 0).text()
            custom_input = self.variants_table.item(row, 1).text() == "Да"
            variants.append({"text": text, "custom_input": custom_input})

        new_template = {"text": text, "variants": variants}

        if category not in self.templates:
            self.templates[category] = []
        self.templates[category].append(new_template)
        self.populate_tree()
        self.clear_editor()
        self.parent().status_bar.showMessage(f"Добавлен шаблон: {text}")

    def edit_template(self):
        """Редактирование существующего шаблона"""
        if not self.selected_template:
            QMessageBox.warning(self, "Ошибка", "Выберите шаблон для редактирования")
            return

        old_category, old_template = self.selected_template
        new_category = self.category_combo.currentText().strip()
        new_text = self.template_text.text().strip()
        if not new_category or not new_text:
            QMessageBox.warning(self, "Ошибка", "Укажите категорию и текст шаблона")
            return

        # Собираем варианты из таблицы
        variants = []
        for row in range(self.variants_table.rowCount()):
            text = self.variants_table.item(row, 0).text()
            custom_input = self.variants_table.item(row, 1).text() == "Да"
            variants.append({"text": text, "custom_input": custom_input})

        new_template = {"text": new_text, "variants": variants}

        # Удаляем старый шаблон
        self.templates[old_category].remove(old_template)
        if not self.templates[old_category]:
            del self.templates[old_category]

        # Добавляем новый
        if new_category not in self.templates:
            self.templates[new_category] = []
        self.templates[new_category].append(new_template)

        self.populate_tree()
        self.clear_editor()
        self.selected_template = None
        self.parent().status_bar.showMessage(f"Шаблон обновлен: {new_text}")

    def delete_template(self):
        """Удаление шаблона"""
        if not self.selected_template:
            QMessageBox.warning(self, "Ошибка", "Выберите шаблон для удаления")
            return

        category, template = self.selected_template
        self.templates[category].remove(template)
        if not self.templates[category]:
            del self.templates[category]
        self.populate_tree()
        self.clear_editor()
        self.selected_template = None
        self.parent().status_bar.showMessage(f"Шаблон удален")

    def clear_editor(self):
        """Очистка полей редактора"""
        self.template_text.clear()
        self.variants_table.setRowCount(0)  # Очищаем таблицу
        self.variant_text.clear()
        self.custom_input_check.setChecked(False)

    def save_and_close(self):
        """Сохранение изменений и закрытие"""
        try:
            with open(self.templates_file, "w", encoding="utf-8") as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=4)
            self.parent().templates = self.templates.copy()
            self.parent().reload_templates()
            self.parent().status_bar.showMessage("Шаблоны сохранены")
            self.accept()
        except Exception as e:
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось сохранить шаблоны: {str(e)}"
            )

    def update_preview(self):
        if not self.selected_template:
            return
        category, template = self.selected_template
        text = self.template_text.text()
        variants = [
            self.variants_table.item(i, 0).text()
            for i in range(self.variants_table.rowCount())
        ]
        preview = f"{text}\n" + "\n".join([f"  - {v}" for v in variants])
        self.preview_text.setPlainText(preview)

    def toggle_custom_input(self, item):
        """Переключение значения 'Да'/'Нет' в колонке 'Пользовательский ввод' по двойному клику"""
        column = item.column()
        row = item.row()

        # Обрабатываем только колонку "Пользовательский ввод" (индекс 1)
        if column == 1:
            current_value = self.variants_table.item(row, 1).text()
            new_value = "Нет" if current_value == "Да" else "Да"
            new_item = QTableWidgetItem(new_value)
            new_item.setFlags(
                new_item.flags() & ~Qt.ItemFlag.ItemIsEditable
            )  # Не редактируемая
            self.variants_table.setItem(row, 1, new_item)
            # Обновляем предпросмотр, если нужно
            self.update_preview()


class ThemeManager:
    DARK_THEME = """
        QMainWindow, QDialog {
            background-color: #1F2526;
        }
        QWidget {
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        QGroupBox {
            font-size: 12px;
            font-weight: bold;
            border: 1px solid #303940;
            border-radius: 5px;
            margin-top: 10px;
            padding: 10px;
            background-color: #2A3033;
            color: #D3D7DA;
        }
        QLabel {
            color: #D3D7DA;
        }
        QLineEdit, QComboBox {
            padding: 6px;
            border: 1px solid #303940;
            border-radius: 4px;
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QLineEdit:focus, QComboBox:focus {
            border-width: 2px;
            border-color: #409EFF;
        }
        QPushButton {
            padding: 8px 16px;
            border: 1px solid #303940;
            border-radius: 4px;
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QPushButton:hover {
            background-color: #3A4446;
            border-color: #409EFF;
            color: #409EFF;
        }
        QPushButton:pressed {
            background-color: #1E2527;
        }
        QTextEdit, QTableWidget, QTreeWidget, QListWidget {
            border: 1px solid #303940;
            border-radius: 4px;
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QTabWidget::pane {
            border: 1px solid #303940;
            background-color: #2A3033;
        }
        QTabBar::tab {
            padding: 10px 20px;
            border-bottom: 2px solid transparent;
            color: #E6ECEF;
            background-color: #2A3033;
        }
        QTabBar::tab:selected {
            border-bottom: 2px solid #409EFF;
            color: #409EFF;
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3A4446, stop:1 #2A3033);
        }
        QTreeWidget::item:selected, QListWidget::item:selected, QTableWidget::item:selected {
            background-color: #3A4446;
            color: #409EFF;
        }
        QToolBar {
            background-color: #2A3033;
            border-bottom: 1px solid #303940;
            padding: 4px;
        }
        QToolButton {
            padding: 6px;
            margin: 2px;
            border-radius: 4px;
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QToolButton:hover {
            background-color: #3A4446;
            color: #409EFF;
        }
        QStatusBar {
            background-color: #2A3033;
            border-top: 1px solid #303940;
            color: #A6ACAF;
        }
        QMenuBar {
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QMenuBar::item {
            background-color: #2A3033;
            color: #E6ECEF;
            padding: 4px 8px;
        }
        QMenuBar::item:selected {
            background-color: #3A4446;
            color: #409EFF;
        }
        QMenu {
            background-color: #2A3033;
            border: 1px solid #303940;
            color: #E6ECEF;
        }
        QMenu::item {
            padding: 4px 20px;
            background-color: #2A3033;
            color: #E6ECEF;
        }
        QMenu::item:selected {
            background-color: #3A4446;
            color: #409EFF;
        }
        QScrollBar:vertical, QScrollBar:horizontal {
            background-color: #2A3033;
            width: 14px;
            height: 14px;
            margin: 0px;
            border: 1px solid #303940;
        }
        QScrollBar::handle {
            background-color: #606266;
            border-radius: 7px;
        }
        QScrollBar::handle:hover {
            background-color: #A6ACAF;
        }
        QScrollBar::add-line, QScrollBar::sub-line {
            background: none;
            border: none;
        }
        QScrollBar::add-page, QScrollBar::sub-page {
            background-color: #2A3033;
        }
        QHeaderView::section {
            background-color: #252B2D;
            color: #E6ECEF;
            padding: 4px;
            border: 1px solid #303940;
        }
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        QComboBox::down-arrow {
            width: 10px;
            height: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #2A3033;
            border: 1px solid #303940;
            color: #E6ECEF;
            selection-background-color: #3A4446;
            selection-color: #409EFF;
        }
        QSplitter::handle {
            background-color: #303940;
            width: 4px;
            height: 4px;
        }
        QSplitter::handle:hover {
            background-color: #409EFF;
        }
        QSplitter::handle:pressed {
            background-color: #1E2527;
        }
    """

    LIGHT_THEME = """
        QMainWindow, QDialog {
            background-color: #F5F6FA;
        }
        QWidget {
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        QGroupBox {
            font-size: 12px;
            font-weight: bold;
            border: 1px solid #DCDFE6;
            border-radius: 5px;
            margin-top: 10px;
            padding: 10px;
            background-color: #FFFFFF;
            color: #212529;
        }
        QLabel {
            color: #212529;
        }
        QLineEdit, QComboBox {
            padding: 6px;
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            background-color: #FFFFFF;
            color: #303133;
        }
        QLineEdit:focus, QComboBox:focus {
            border-width: 2px;
            border-color: #409EFF;
        }
        QPushButton {
            padding: 8px 16px;
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            background-color: #FFFFFF;
            color: #606266;
        }
        QPushButton:hover {
            background-color: #ECF5FF;
            border-color: #409EFF;
            color: #409EFF;
        }
        QPushButton:pressed {
            background-color: #D6EBFF;
        }
        QTextEdit, QTableWidget, QTreeWidget, QListWidget {
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            background-color: #FFFFFF;
            color: #303133;
        }
        QTabWidget::pane {
            border: 1px solid #DCDFE6;
            background-color: #FFFFFF;
        }
        QTabBar::tab {
            padding: 10px 20px;
            border-bottom: 2px solid transparent;
            color: #606266;
            background-color: #FFFFFF;
        }
        QTabBar::tab:selected {
            border-bottom: 2px solid #409EFF;
            color: #409EFF;
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ECF5FF, stop:1 #FFFFFF);
        }
        QTreeWidget::item:selected, QListWidget::item:selected, QTableWidget::item:selected {
            background-color: #E6F7FF;
            color: #409EFF;
        }
        QToolBar {
            background-color: #FFFFFF;
            border-bottom: 1px solid #DCDFE6;
            padding: 4px;
        }
        QToolButton {
            padding: 6px;
            margin: 2px;
            border-radius: 4px;
            background-color: #FFFFFF;
            color: #606266;
        }
        QToolButton:hover {
            background-color: #ECF5FF;
            color: #409EFF;
        }
        QStatusBar {
            background-color: #FFFFFF;
            border-top: 1px solid #DCDFE6;
            color: #606266;
        }
        QMenuBar {
            background-color: #FFFFFF;
            color: #303133;
        }
        QMenuBar::item {
            background-color: #FFFFFF;
            color: #303133;
            padding: 4px 8px;
        }
        QMenuBar::item:selected {
            background-color: #ECF5FF;
            color: #409EFF;
        }
        QMenu {
            background-color: #FFFFFF;
            border: 1px solid #DCDFE6;
            color: #303133;
        }
        QMenu::item {
            padding: 4px 20px;
            background-color: #FFFFFF;
            color: #303133;
        }
        QMenu::item:selected {
            background-color: #ECF5FF;
            color: #409EFF;
        }
        QScrollBar:vertical, QScrollBar:horizontal {
            background-color: #FFFFFF;
            width: 14px;
            height: 14px;
            margin: 0px;
            border: 1px solid #DCDFE6;
        }
        QScrollBar::handle {
            background-color: #C0C4CC;
            border-radius: 7px;
        }
        QScrollBar::handle:hover {
            background: #A6ACAF;
        }
        QScrollBar::add-line, QScrollBar::sub-line {
            background: none;
            border: none;
        }
        QScrollBar::add-page, QScrollBar::sub-page {
            background-color: #FFFFFF;
        }
        QHeaderView::section {
            background-color: #F5F6FA;
            color: #303133;
            padding: 4px;
            border: 1px solid #DCDFE6;
        }
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        QComboBox::down-arrow {
            width: 10px;
            height: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #FFFFFF;
            border: 1px solid #DCDFE6;
            color: #303133;
            selection-background-color: #ECF5FF;
            selection-color: #409EFF;
        }
        QSplitter::handle {
            background-color: #DCDFE6;
            width: 4px;
            height: 4px;
            border-radius: 3px
        }
        QSplitter::handle:hover {
            background-color: #409EFF;
        }
        QSplitter::handle:pressed {
            background-color: #D6EBFF;
        }
    """

    @staticmethod
    def apply_theme(widget, dark_mode):
        """Применение темы к виджету"""
        if dark_mode:
            widget.setStyleSheet(ThemeManager.DARK_THEME)
        else:
            widget.setStyleSheet(ThemeManager.LIGHT_THEME)


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
