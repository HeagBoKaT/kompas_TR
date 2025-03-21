import os
import sys
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, font
from tkinter import filedialog
import pythoncom
import win32com.client
from win32com.client import Dispatch, gencache
import re
import gc

class KompasApp:
    def __init__(self, root):
        """Инициализация приложения"""
        self.root = root
        self.root.title("Редактор технических требований KOMPAS-3D")
        self.root.geometry("1400x900")
        self.root.minsize(1000, 700)
        
        # Установка иконки приложения
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        
        # Инициализация переменных
        self.templates = {}
        self.template_search_var = tk.StringVar()
        self.template_search_var.trace("w", self.filter_templates)
        
        # Переменная для автонумерации
        self.auto_numbering_var = tk.BooleanVar(value=False)
        self.auto_numbering_var.trace("w", self.toggle_auto_numbering)
        
        # Создаем строку статуса сразу при инициализации
        self.status_bar = ttk.Label(self.root, text="Инициализация...", anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Установка стиля
        self.set_style()
        
        # Загрузка шаблонов
        self.load_templates()
        
        # Создание пользовательского интерфейса
        self.create_ui()
        
        # Создание меню
        self.create_menu()
        
        # Создание горячих клавиш
        self.create_shortcuts()
        
        # Обновление строки статуса
        self.set_status("Готово")
        
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
        self.root.after(1000, self.periodic_update)
        
        # Обработчик закрытия приложения
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def set_style(self):
        """Установка стиля для приложения"""
        style = ttk.Style()
        
        # Настройка глобальных стилей
        style.configure('TFrame', background='#f5f5f5')
        style.configure('TLabel', background='#f5f5f5')
        style.configure('TLabelframe', background='#f5f5f5')
        style.configure('TLabelframe.Label', background='#f5f5f5', font=('Segoe UI', 10, 'bold'))
        
        # Стиль для панели инструментов
        style.configure('Toolbar.TFrame', background='#e9e9e9')
        
        # Стиль для панели поиска
        style.configure('Search.TFrame', background='#e9e9e9')
        
        # Стиль для кнопок
        style.configure('TButton', padding=3)
        
        # Стиль для статусной строки
        style.configure('Status.TLabel', background='#e9e9e9', relief='sunken', anchor='w', padding=(5, 2))
        style.configure('StatusGreen.TLabel', background='#e9e9e9', foreground='green', relief='sunken')
        style.configure('StatusRed.TLabel', background='#e9e9e9', foreground='red', relief='sunken')
        
        # Стиль для заголовков Treeview
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Стиль для вкладок
        style.configure('TNotebook.Tab', padding=[12, 4])
        
        # Настройка конкретных элементов
        style.configure('TButton', padding=5, font=('Segoe UI', 10))
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('TNotebook', background='#f0f0f0')
        style.configure('TNotebook.Tab', padding=[12, 6], font=('Segoe UI', 10))
        style.configure('Treeview', font=('Segoe UI', 10), rowheight=28)
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Настройка для строки статуса
        style.configure('StatusBar.TLabel', padding=3, background='#e0e0e0', relief='sunken')
        
        # Настройка для панели поиска
        style.configure('Search.TEntry', padding=5)
        style.configure('Search.TButton', padding=3)
        
        # Настройка для панели с инструментами
        style.configure('Toolbar.TButton', padding=3)
        
        # Настройка для Listbox (шаблоны)
        self.root.option_add('*Listbox*font', ('Segoe UI', 10))
        self.root.option_add('*Listbox*background', '#ffffff')
        self.root.option_add('*Listbox*selectBackground', '#4a6984')
        self.root.option_add('*Listbox*selectForeground', '#ffffff')
        
    def load_templates(self):
        """Загрузка шаблонов технических требований из файла JSON"""
        try:
            # Определяем путь к файлу шаблонов в корневой папке пользователя
            user_home = os.path.expanduser("~")
            app_folder = os.path.join(user_home, "KOMPAS-TR")
            
            # Создаем папку для приложения, если она не существует
            if not os.path.exists(app_folder):
                os.makedirs(app_folder)
                
            self.templates_file = os.path.join(app_folder, "templates.json")
            
            # Проверяем существование файла
            if not os.path.exists(self.templates_file):
                self.set_status("Файл шаблонов не найден, создаем новый")
                
                # Проверяем, есть ли файл шаблонов в директории приложения (для обратной совместимости)
                old_templates_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates.json")
                
                if os.path.exists(old_templates_file):
                    # Копируем существующий файл шаблонов из директории приложения
                    with open(old_templates_file, 'r', encoding='utf-8') as f_old:
                        templates_data = json.load(f_old)
                    
                    with open(self.templates_file, 'w', encoding='utf-8') as f_new:
                        json.dump(templates_data, f_new, ensure_ascii=False, indent=4)
                        
                    self.set_status("Файл шаблонов перенесен в папку пользователя")
                else:
                    # Создаем пустой файл шаблонов
                    with open(self.templates_file, 'w', encoding='utf-8') as f:
                        json.dump({"Общие": []}, f, ensure_ascii=False, indent=4)
            
            # Загружаем шаблоны из файла
            with open(self.templates_file, 'r', encoding='utf-8') as f:
                self.templates = json.load(f)
                
            self.set_status(f"Загружено {sum(len(templates) for templates in self.templates.values())} шаблонов")
            
        except Exception as e:
            self.set_status(f"Ошибка загрузки шаблонов: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблоны: {str(e)}")
            # Создаем пустой словарь шаблонов в случае ошибки
            self.templates = {"Общие": []}
            
    def connect_to_kompas(self):
        """Подключение к KOMPAS-3D"""
        try:
            # Проверяем, есть ли уже подключение
            if hasattr(self, 'app7') and self.app7:
                # Проверяем, работает ли подключение
                try:
                    app_name = self.app7.ApplicationName(FullName=False)
                    self.connect_status.config(text="🟢 Подключено", foreground='green')
                    self.set_status(f"Уже подключено к {app_name}")
                    return True
                except Exception as e:
                    # Подключение не работает, пробуем заново
                    self.app7 = None
                    self.set_status("Ошибка подключения, пробуем переподключиться...")
            
            # Пробуем подключиться к запущенному KOMPAS-3D
            try:
                self.set_status("Попытка подключения к запущенному KOMPAS-3D...")
                self.app7 = win32com.client.Dispatch("Kompas.Application.7")
                app_name = self.app7.ApplicationName(FullName=False)
                self.module7, self.api7, self.const7 = self.get_kompas_api7()
                self.connect_status.config(text="🟢 Подключено", foreground='green')
                self.set_status(f"Подключено к запущенному {app_name}")
                
                # Обновляем дерево документов
                self.update_documents_tree()
                
                return True
            except Exception as e:
                # Если не удалось подключиться к запущенному, пробуем запустить новый
                try:
                    self.set_status("Попытка запуска KOMPAS-3D...")
                    self.app7 = win32com.client.Dispatch("Kompas.Application.7")
                    self.app7.Visible = True
                    self.app7.HideMessage = True
                    self.module7, self.api7, self.const7 = self.get_kompas_api7()
                    app_name = self.app7.ApplicationName(FullName=False)
                    self.connect_status.config(text="🟢 Подключено", foreground='green')
                    self.set_status(f"Запущен и подключен {app_name}")
                    
                    # Обновляем дерево документов
                    self.update_documents_tree()
                    
                    return True
                except Exception as e:
                    self.connect_status.config(text="🔴 Нет подключения", foreground='red')
                    error_message = self.handle_kompas_error(e, "подключения")
                    self.set_status("Не удалось подключиться к KOMPAS-3D")
                    messagebox.showerror("Ошибка подключения", error_message)
                    return False
                    
        except Exception as e:
            self.connect_status.config(text="🔴 Нет подключения", foreground='red')
            error_message = self.handle_kompas_error(e, "подключения")
            self.set_status("Ошибка при подключении к KOMPAS-3D")
            messagebox.showerror("Ошибка подключения", error_message)
            return False
            
    def check_kompas_connection(self):
        """Проверка подключения к KOMPAS-3D с выводом сообщения"""
        if self.is_kompas_running():
            app_name = self.app7.ApplicationName(FullName=True)
            version = self.app7.ApplicationVersion()
            messagebox.showinfo("Информация о подключении", 
                              f"Подключено к KOMPAS-3D\n\n"
                              f"Приложение: {app_name}\n"
                              f"Версия: {version}")
            self.set_status(f"Подключено к {app_name} версии {version}")
            return True
        else:
            result = messagebox.askyesno("Нет подключения", 
                                       "Нет подключения к KOMPAS-3D.\n\n"
                                       "Хотите попробовать подключиться?")
            if result:
                return self.connect_to_kompas()
            return False
            
    def get_kompas_api7(self):
        """Получение объектов API Kompas 3D версии 7"""
        module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        api = module.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                    pythoncom.IID_IDispatch))
        const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        return module, api, const
        
    def is_kompas_running(self):
        """Проверка подключения к KOMPAS-3D"""
        try:
            return hasattr(self, 'app7') and self.app7 is not None
        except:
            return False

    def create_ui(self):
        """Создание пользовательского интерфейса"""
        # Создание меню
        self.create_menu()
        
        # Главный фрейм
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Верхняя панель - панель инструментов
        toolbar_frame = self.create_toolbar(main_frame)
        toolbar_frame.pack(side=tk.TOP, fill=tk.X)
        
        # Верхняя панель - информация о документе
        doc_frame = ttk.LabelFrame(main_frame, text="Активный документ")
        doc_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Создание информационной строки для активного документа
        doc_info_frame = ttk.Frame(doc_frame)
        doc_info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Иконка документа и статус соединения
        icon_frame = ttk.Frame(doc_info_frame)
        icon_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        # Иконка KOMPAS
        ttk.Label(icon_frame, text="📐", font=('Segoe UI', 16)).pack(side=tk.TOP)
        
        # Статус подключения
        self.connect_status = ttk.Label(icon_frame, text="🔴 Нет подключения", foreground='red')
        self.connect_status.pack(side=tk.TOP)
        
        # Информация о документе
        doc_text_frame = ttk.Frame(doc_info_frame)
        doc_text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.active_doc_label = ttk.Label(doc_text_frame, text="Нет активного документа", 
                                       font=('Segoe UI', 10, 'bold'), wraplength=600)
        self.active_doc_label.pack(anchor="w")
        
        doc_desc_label = ttk.Label(doc_text_frame, 
                                 text="Выберите документ из списка или откройте новый в KOMPAS-3D")
        doc_desc_label.pack(anchor="w")
        
        # Разделение на 2 панели
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Левая панель - дерево документов
        left_frame = ttk.LabelFrame(paned, text="Открытые документы")
        
        # Панель поиска для документов
        search_doc_frame = ttk.Frame(left_frame, style='Search.TFrame')
        search_doc_frame.pack(fill=tk.X, padx=5, pady=(5, 0))
        
        ttk.Label(search_doc_frame, text="🔍", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.doc_search_var = tk.StringVar()
        self.doc_search_var.trace_add("write", self.filter_documents_tree)
        self.doc_search_entry = ttk.Entry(search_doc_frame, textvariable=self.doc_search_var)
        self.doc_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Кнопка обновления дерева документов
        refresh_btn = ttk.Button(search_doc_frame, text="🔄", width=3,
                               command=self.update_documents_tree)
        refresh_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Создание и настройка Treeview для отображения документов
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Улучшенный Treeview с скроллбарами для вертикальной и горизонтальной прокрутки
        self.doc_tree = ttk.Treeview(tree_frame, style='Treeview')
        self.doc_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Вертикальная прокрутка
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.doc_tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.doc_tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Горизонтальная прокрутка
        h_scrollbar = ttk.Scrollbar(left_frame, orient="horizontal", command=self.doc_tree.xview)
        h_scrollbar.pack(fill=tk.X)
        self.doc_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Настройка колонок
        self.doc_tree["columns"] = ("type", "path")
        self.doc_tree.column("#0", width=150, minwidth=100)
        self.doc_tree.column("type", width=100, minwidth=50)
        self.doc_tree.column("path", width=300, minwidth=200)
        
        self.doc_tree.heading("#0", text="Имя")
        self.doc_tree.heading("type", text="Тип")
        self.doc_tree.heading("path", text="Путь")
        
        # Обработчик событий для дерева документов
        self.doc_tree.bind("<Double-1>", self.on_document_double_click)
        
        # Контекстное меню для дерева документов
        self.doc_tree_context_menu = tk.Menu(self.doc_tree, tearoff=0)
        self.doc_tree_context_menu.add_command(label="Активировать", command=self.activate_selected_document)
        self.doc_tree_context_menu.add_command(label="Показать информацию", command=self.show_document_info)
        self.doc_tree.bind("<Button-3>", self.show_doc_tree_context_menu)
        
        paned.add(left_frame, weight=1)
        
        # Правая панель - шаблоны и редактирование технических требований
        right_frame = ttk.Frame(paned)
        
        # Разделение правой панели по вертикали
        right_paned = ttk.PanedWindow(right_frame, orient=tk.VERTICAL)
        right_paned.pack(fill=tk.BOTH, expand=True)
        
        # Блок с шаблонами технических требований
        templates_frame = ttk.LabelFrame(right_paned, text="Шаблоны технических требований")
        
        # Панель поиска для шаблонов
        search_template_frame = ttk.Frame(templates_frame)
        search_template_frame.pack(fill=tk.X, padx=5, pady=5)
        
        search_label = ttk.Label(search_template_frame, text="Поиск шаблона:")
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.template_search_var = tk.StringVar()
        self.template_search_var.trace_add("write", self.filter_templates)
        self.template_search_entry = ttk.Entry(search_template_frame, textvariable=self.template_search_var)
        self.template_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Создание вкладок для категорий шаблонов
        self.template_tabs = ttk.Notebook(templates_frame)
        self.template_tabs.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Заполнение вкладок
        self.populate_template_tabs()
        
        right_paned.add(templates_frame, weight=2)
        
        # Блок с текущими техническими требованиями
        current_reqs_frame = ttk.LabelFrame(right_paned, text="Текущие технические требования")
        
        # Текстовое поле для отображения/редактирования технических требований
        text_frame = ttk.Frame(current_reqs_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создание и настройка Text widget с возможностью вертикальной прокрутки
        self.current_reqs_text = tk.Text(text_frame, wrap=tk.WORD, undo=True)
        self.current_reqs_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Настройка стилей текста
        self.current_reqs_text.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))
        self.current_reqs_text.tag_configure("italic", font=("TkDefaultFont", 10, "italic"))
        self.current_reqs_text.tag_configure("underline", underline=1)
        
        # Vertical scrollbar
        v_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.current_reqs_text.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.current_reqs_text.configure(yscrollcommand=v_scrollbar.set)
        
        # Кнопки управления техническими требованиями
        buttons_frame = ttk.Frame(current_reqs_frame)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(buttons_frame, text="Получить", 
                  command=self.get_technical_requirements).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Сохранить", 
                  command=self.save_technical_requirements).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Применить", 
                  command=lambda: self.apply_technical_requirements()).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Очистить", 
                  command=lambda: self.current_reqs_text.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=5)
        
        right_paned.add(current_reqs_frame, weight=3)
        
        paned.add(right_frame, weight=3)
        
        # Создание строки статуса
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Разделитель
        ttk.Separator(status_frame, orient='horizontal').pack(fill=tk.X)
        
        # Создание фрейма для строки статуса
        status_inner_frame = ttk.Frame(status_frame, style='Status.TFrame')
        status_inner_frame.pack(fill=tk.X, padx=1, pady=1)
        
        # Удаляем старую строку статуса
        if hasattr(self, 'status_bar'):
            self.status_bar.destroy()
        
        # Левая часть строки статуса - сообщения
        self.status_bar = ttk.Label(status_inner_frame, text="Готово", style='Status.TLabel')
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=2)
        
        # Правая часть строки статуса - дополнительная информация
        status_right_frame = ttk.Frame(status_inner_frame)
        status_right_frame.pack(side=tk.RIGHT, padx=5)
        
        # Количество открытых документов
        self.docs_count_label = ttk.Label(status_right_frame, text="Документов: 0")
        self.docs_count_label.pack(side=tk.RIGHT, padx=5)
        
        # Версия приложения
        version_label = ttk.Label(status_right_frame, text="v1.0 (2025)")
        version_label.pack(side=tk.RIGHT, padx=5)
        
    def create_menu(self):
        """Создание главного меню"""
        menu = tk.Menu(self.root)
        self.root.config(menu=menu)
        
        # Меню "Файл"
        file_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Подключиться к KOMPAS-3D", 
                             command=self.connect_to_kompas,
                             accelerator="Ctrl+K")
        file_menu.add_command(label="Проверить подключение", 
                             command=self.check_kompas_connection)
        file_menu.add_separator()
        file_menu.add_command(label="Получить технические требования", 
                              command=self.get_technical_requirements,
                              accelerator="Ctrl+G")
        file_menu.add_command(label="Сохранить технические требования", 
                              command=self.save_technical_requirements,
                              accelerator="Ctrl+S")
        file_menu.add_command(label="Применить технические требования", 
                              command=lambda: self.apply_technical_requirements(),
                              accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Отключиться от KOMPAS-3D", 
                              command=self.disconnect_from_kompas)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.on_closing, accelerator="Alt+F4")
        
        # Меню "Инструменты"
        tools_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Инструменты", menu=tools_menu)
        tools_menu.add_command(label="Редактировать файл шаблонов", 
                               command=self.edit_templates_file)
        tools_menu.add_command(label="Обновить шаблоны", 
                               command=self.reload_templates,
                               accelerator="F5")
        tools_menu.add_separator()
        tools_menu.add_command(label="Обновить список документов", 
                               command=lambda: self.update_documents_tree(),
                               accelerator="F6")
        # Меню "Помощь"
        help_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Помощь", menu=help_menu)
        help_menu.add_command(label="О программе", command=self.show_about)
        help_menu.add_command(label="Горячие клавиши", command=self.show_shortcuts)

    def create_toolbar(self, parent):
        """Создание панели инструментов"""
        toolbar_frame = ttk.Frame(parent, style='Toolbar.TFrame')
        toolbar_frame.pack(side=tk.TOP, fill=tk.X)
        
        # Разделитель групп
        def add_separator():
            ttk.Separator(toolbar_frame, orient='vertical').pack(side=tk.LEFT, padx=5, fill=tk.Y, pady=2)
        
        # Группа 1: Подключение к KOMPAS
        # Кнопка подключения к Kompas
        connect_btn = ttk.Button(toolbar_frame, text="🔌", width=3, 
                              command=self.connect_to_kompas)
        connect_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(connect_btn, "Подключиться к KOMPAS-3D (Ctrl+K)")
        
        # Кнопка отключения от KOMPAS-3D
        disconnect_btn = ttk.Button(toolbar_frame, text="🚫", width=3, 
                                  command=self.disconnect_from_kompas)
        disconnect_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(disconnect_btn, "Отключиться от KOMPAS-3D")
        
        # Кнопка проверки подключения
        check_connect_btn = ttk.Button(toolbar_frame, text="🔍", width=3, 
                                     command=self.check_kompas_connection)
        check_connect_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(check_connect_btn, "Проверить подключение к KOMPAS-3D")
        
        # Кнопка обновления списка документов
        refresh_btn = ttk.Button(toolbar_frame, text="🔄", width=3, 
                              command=lambda: self.update_documents_tree())
        refresh_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(refresh_btn, "Обновить список документов (F6)")
        
        add_separator()
        
        # Группа 2: Работа с техническими требованиями
        # Кнопка получения технических требований
        get_btn = ttk.Button(toolbar_frame, text="📥", width=3, 
                          command=self.get_technical_requirements)
        get_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(get_btn, "Получить технические требования (Ctrl+G)")
        
        # Кнопка сохранения технических требований
        save_btn = ttk.Button(toolbar_frame, text="💾", width=3, 
                           command=self.save_technical_requirements)
        save_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(save_btn, "Сохранить технические требования (Ctrl+S)")
        
        # Кнопка применения технических требований
        apply_btn = ttk.Button(toolbar_frame, text="🔄", width=3, 
                              command=lambda: self.apply_technical_requirements())
        apply_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(apply_btn, "Применить технические требования (Ctrl+E)")
        
        add_separator()
 
        # Группа 3: Работа с шаблонами
        # Кнопка редактирования шаблонов
        edit_templates_btn = ttk.Button(toolbar_frame, text="📝", width=3, 
                                    command=self.edit_templates_file)
        edit_templates_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(edit_templates_btn, "Редактировать файл шаблонов")
        
        # Кнопка обновления шаблонов
        reload_templates_btn = ttk.Button(toolbar_frame, text="📋", width=3, 
                                      command=self.reload_templates)
        reload_templates_btn.pack(side=tk.LEFT, padx=2)
        self.create_tooltip(reload_templates_btn, "Обновить шаблоны (F5)")
        
        
        return toolbar_frame
        
    def create_tooltip(self, widget, text):
        """Создание всплывающей подсказки для виджета"""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Создание окна подсказки
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            label = ttk.Label(self.tooltip, text=text, background="#ffffe0", 
                           relief="solid", borderwidth=1, padding=(5, 2))
            label.pack()
            
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
                
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
        
    def create_shortcuts(self):
        """Создание горячих клавиш"""
        self.root.bind("<Control-k>", lambda event: self.connect_to_kompas())
        self.root.bind("<Control-g>", lambda event: self.get_technical_requirements())
        self.root.bind("<Control-s>", lambda event: self.save_technical_requirements())
        self.root.bind("<Control-e>", lambda event: self.apply_technical_requirements())
        self.root.bind("<F5>", lambda event: self.reload_templates())
        self.root.bind("<F6>", lambda event: self.update_documents_tree())
        self.root.bind("<Control-f>", lambda event: self.focus_search())

    def focus_search(self):
        """Установка фокуса на поле поиска"""
        current_tab = self.template_tabs.index(self.template_tabs.select())
        if current_tab == 0:  # Если активна первая вкладка
            self.template_search_entry.focus_set()
        else:
            self.doc_search_entry.focus_set()
            
    def filter_documents_tree(self, *args):
        """Фильтрация дерева документов по поисковому запросу"""
        search_term = self.doc_search_var.get().lower()
        self.update_documents_tree(search_term)
        
    def filter_templates(self, *args):
        """Фильтрация шаблонов по поисковому запросу"""
        search_term = self.template_search_var.get().lower()
        
        # Заполнение вкладок с учетом поискового запроса
        self.populate_template_tabs(search_term)
        
        # Обновление строки статуса
        if search_term:
            # Подсчет общего количества найденных шаблонов
            found_count = 0
            for category, items in self.templates.items():
                for item in items:
                    if search_term in item.lower() or search_term in category.lower():
                        found_count += 1
            
            if found_count == 0:
                self.set_status(f"По запросу '{search_term}' шаблонов не найдено")
            else:
                self.set_status(f"Найдено шаблонов: {found_count} по запросу '{search_term}'")
        else:
            self.set_status("Показаны все шаблоны")
            
    def show_doc_tree_context_menu(self, event):
        """Показать контекстное меню для дерева документов"""
        if self.doc_tree.identify_row(event.y):
            self.doc_tree.selection_set(self.doc_tree.identify_row(event.y))
            self.doc_tree_context_menu.post(event.x_root, event.y_root)
            
    def activate_selected_document(self):
        """Активация выбранного документа в дереве"""
        selected = self.doc_tree.selection()
        if not selected:
            self.set_status("Нет выбранного документа")
            return False
            
        try:
            if not hasattr(self, 'app7') or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, 'app7') or not self.app7:
                    return
                    
            # Получаем имя выбранного документа
            doc_name = self.doc_tree.item(selected[0], 'text')
            
            # Ищем документ в списке открытых документов
            documents = self.app7.Documents
            for i in range(documents.Count):
                doc = documents.Item(i)
                if doc.Name == doc_name:
                    # Активируем документ
                    doc.Active = True
                    self.update_active_document_info()
                    self.set_status(f"Документ {doc_name} активирован")
                    return True
                    
            self.set_status(f"Документ {doc_name} не найден в списке открытых документов")
            return False
            
        except Exception as e:
            error_message = self.handle_kompas_error(e, "активации документа")
            self.set_status("Ошибка при активации документа")
            messagebox.showerror("Ошибка", error_message)
            return False
            
    def show_document_info(self):
        """Показать подробную информацию о выбранном документе"""
        selected = self.doc_tree.selection()
        if selected:
            item = selected[0]
            doc_name = self.doc_tree.item(item, "text")
            doc_type = self.doc_tree.item(item, "values")[0]
            doc_path = self.doc_tree.item(item, "values")[1]
            
            info = f"Имя: {doc_name}\nТип: {doc_type}\nПуть: {doc_path}"
            
            messagebox.showinfo("Информация о документе", info)
            
    def edit_templates_file(self):
        """Открытие файла шаблонов во внешнем редакторе"""
        try:
            if not os.path.exists(self.templates_file):
                self.set_status("Файл шаблонов не найден, создаем новый")
                with open(self.templates_file, 'w', encoding='utf-8') as f:
                    json.dump({"Общие": []}, f, ensure_ascii=False, indent=4)
                    
            # Открываем файл в редакторе по умолчанию
            os.startfile(self.templates_file)
            self.set_status(f"Файл шаблонов открыт для редактирования: {self.templates_file}")
            
            # Запрашиваем у пользователя, нужно ли обновить шаблоны после редактирования
            result = messagebox.askyesno("Обновление шаблонов", 
                                       "После завершения редактирования файла шаблонов, "
                                       "хотите ли вы обновить шаблоны в программе?")
            if result:
                # Планируем обновление шаблонов через некоторое время
                self.root.after(1000, self.reload_templates)
                
        except Exception as e:
            self.set_status(f"Ошибка при открытии файла шаблонов: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось открыть файл шаблонов: {str(e)}")
            
    def reload_templates(self):
        """Перезагрузка шаблонов из файла"""
        try:
            # Сохраняем текущий поисковый запрос
            current_search = self.template_search_var.get()
            
            # Загружаем шаблоны
            self.load_templates()
            
            # Обновляем вкладки с шаблонами
            self.populate_template_tabs()
            
            # Восстанавливаем поиск, если был
            if current_search:
                self.template_search_var.set(current_search)
                self.filter_templates()
                
            self.set_status("Шаблоны успешно обновлены")
            
        except Exception as e:
            self.set_status(f"Ошибка при обновлении шаблонов: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось обновить шаблоны: {str(e)}")
            
    def show_about(self):
        """Показать информацию о программе"""
        about_text = """
        Редактор технических требований для KOMPAS-3D
        
        Программа для редактирования и вставки типовых 
        текстов в технические требования чертежей KOMPAS-3D.
        
        2025
        """
        messagebox.showinfo("О программе", about_text)
        
    def show_shortcuts(self):
        """Показать горячие клавиши"""
        shortcuts_text = """
        Горячие клавиши:
        Ctrl+K - Подключиться к KOMPAS-3D
        Ctrl+G - Получить технические требования
        Ctrl+S - Сохранить технические требования
        Ctrl+E - Применить технические требования
        F5 - Обновить шаблоны
        F6 - Обновить список документов
        """
        messagebox.showinfo("Горячие клавиши", shortcuts_text)
        
    def set_status(self, text):
        """Установка текста в строке статуса"""
        self.status_bar.config(text=text)
        
    def populate_template_tabs(self, search_term=None):
        """Заполнение вкладок шаблонами технических требований"""
        # Очистка существующих вкладок
        for tab in self.template_tabs.tabs():
            self.template_tabs.forget(tab)
            
        # Создание вкладки "Все" для общего поиска
        all_tab = ttk.Frame(self.template_tabs)
        self.template_tabs.add(all_tab, text="Все")
        
        # Создаем фрейм для Listbox и скроллбара
        all_list_frame = ttk.Frame(all_tab)
        all_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Listbox для всех шаблонов при поиске
        all_listbox = tk.Listbox(all_list_frame, font=('Segoe UI', 10), activestyle='dotbox', 
                                 selectbackground='#4a6984', selectforeground='white')
        all_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Добавление прокрутки для вкладки "Все"
        all_scrollbar = ttk.Scrollbar(all_list_frame, orient="vertical", command=all_listbox.yview)
        all_listbox.configure(yscrollcommand=all_scrollbar.set)
        all_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Обработчик для вставки шаблона с общей вкладки
        all_listbox.bind("<Double-1>", lambda event, lb=all_listbox: 
                        self.insert_template(lb.get(lb.curselection()) if lb.curselection() else ""))
        
        # Счетчик найденных шаблонов
        found_count = 0
            
        # Создание вкладок для каждой категории шаблонов
        for category, templates in self.templates.items():
            tab = ttk.Frame(self.template_tabs)
            self.template_tabs.add(tab, text=category)
            
            # Создаем фрейм для Listbox и скроллбара
            list_frame = ttk.Frame(tab)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Создание и настройка Listbox для шаблонов
            templates_listbox = tk.Listbox(list_frame, font=('Segoe UI', 10), activestyle='dotbox', 
                                          selectbackground='#4a6984', selectforeground='white')
            templates_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Добавление прокрутки
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=templates_listbox.yview)
            templates_listbox.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Заполнение Listbox шаблонами с учетом поиска
            category_found_count = 0
            if search_term is None or not search_term:
                # Показываем все шаблоны категории если нет поиска
                for template in templates:
                    templates_listbox.insert(tk.END, template)
                    category_found_count += 1
            else:
                # Фильтрация по поиску
                for template in templates:
                    if search_term in template.lower() or search_term in category.lower():
                        templates_listbox.insert(tk.END, template)
                        # Добавляем шаблон и во вкладку "Все"
                        all_listbox.insert(tk.END, f"[{category}] {template}")
                        category_found_count += 1
                        found_count += 1
            
            # Обработчик для вставки выбранного шаблона
            templates_listbox.bind("<Double-1>", lambda event, lb=templates_listbox: 
                                 self.insert_template(lb.get(lb.curselection()) if lb.curselection() else ""))
            
        # Если идет поиск, активируем вкладку "Все"
        if search_term is not None and search_term:
            self.template_tabs.select(0)  # Выбираем первую вкладку ("Все")
            
    def update_active_document_info(self):
        """Обновление информации об активном документе"""
        try:
            if not hasattr(self, 'app7') or not self.app7:
                self.connect_status.config(text="🔴 Нет подключения", foreground='red')
                self.active_doc_label.config(text="Нет активного документа")
                self.set_status("Нет подключения к KOMPAS-3D")
                return
                
            # Получаем активный документ
            active_doc = self.app7.ActiveDocument
            
            if active_doc:
                try:
                    # Получаем имя документа
                    doc_name = active_doc.Name
                    
                    # Определяем тип документа
                    doc_type = "Неизвестный тип"
                    try:
                        # Проверяем, является ли документ чертежом
                        doc2D_s = active_doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['IDrawingDocument'], 
                                                                  pythoncom.IID_IDispatch)
                        doc_type = "Чертеж"
                    except:
                        try:
                            # Проверяем, является ли документ 3D-моделью
                            doc3D_s = active_doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['IDocument3D'], 
                                                                      pythoncom.IID_IDispatch)
                            doc_type = "3D-модель"
                        except:
                            try:
                                # Проверяем, является ли документ спецификацией
                                spec_s = active_doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['ISpecificationDocument'], 
                                                                         pythoncom.IID_IDispatch)
                                doc_type = "Спецификация"
                            except:
                                pass
                    
                    # Получаем путь к документу
                    doc_path = active_doc.Path
                    if doc_path:
                        doc_path = os.path.join(doc_path, doc_name)
                    else:
                        doc_path = "Документ не сохранен"
                    
                    # Обновляем информацию в интерфейсе
                    self.active_doc_label.config(text=f"Документ: {doc_name} ({doc_type})")
                    self.connect_status.config(text="🟢 Подключено", foreground='green')
                    
                    # Обновляем строку статуса
                    self.set_status(f"Активный документ: {doc_name} ({doc_type}) - {doc_path}")
                    
                    # Выделяем активный документ в дереве
                    self.select_document_in_tree(active_doc)
                    
                except Exception as e:
                    self.set_status(f"Ошибка при получении информации о документе: {str(e)}")
                    self.active_doc_label.config(text="Ошибка получения информации о документе")
            else:
                self.active_doc_label.config(text="Нет активного документа")
                self.set_status("Нет активного документа в KOMPAS-3D")
                
        except Exception as e:
            self.set_status(f"Ошибка при обновлении информации о документе: {str(e)}")
            self.active_doc_label.config(text="Ошибка обновления информации")
            
    def on_document_double_click(self, event):
        """Обработка двойного клика на документе в дереве"""
        # Получаем выбранный элемент
        item_id = self.doc_tree.identify('item', event.x, event.y)
        if not item_id:
            return
            
        # Получаем информацию о документе
        doc_type = self.doc_tree.item(item_id, "values")[0]
        
        # Активируем документ
        self.activate_selected_document()
        
        # Если это чертеж, загружаем технические требования
        if doc_type == "Чертеж":
            # Небольшая задержка для завершения активации документа
            self.root.after(500, self.get_technical_requirements)
            self.set_status("Загрузка технических требований...")
            
    def insert_template(self, template_text):
        """Вставка выбранного шаблона в текстовое поле"""
        if template_text:
            # Удаляем название категории в квадратных скобках, если оно есть
            if template_text.startswith('['):
                template_text = template_text[template_text.find(']') + 1:].strip()
                
            cursor_pos = self.current_reqs_text.index(tk.INSERT)
            self.current_reqs_text.insert(cursor_pos, template_text + "\n")
            self.set_status(f"Вставлен шаблон: {template_text[:30]}...")
            
    def get_technical_requirements(self):
        """Получение технических требований из активного документа"""
        try:
            if not hasattr(self, 'module7') or not self.module7:
                self.connect_to_kompas()
                if not hasattr(self, 'module7') or not self.module7:
                    return
            
            # Получаем активный документ
            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.set_status("Нет активного документа")
                messagebox.showwarning("Внимание", "Нет активного документа в КОМПАС-3D")
                return
                
            # Проверяем, является ли документ чертежом
            try:
                # Пробуем получить через интерфейс IDrawingDocument
                try:
                    # Получаем интерфейс чертежа
                    drawing_document = self.module7.IDrawingDocument(active_doc)
                    
                    # Получаем интерфейс технических требований
                    tech_demand = drawing_document.TechnicalDemand
                    
                    # Проверяем, созданы ли технические требования
                    if not tech_demand.IsCreated:
                        self.set_status("В документе отсутствуют технические требования!")
                        messagebox.showwarning("Внимание", "В документе отсутствуют технические требования!")
                        return
                    
                    # Получаем текст технических требований
                    text = tech_demand.Text
                    
                    # Если нет строк, выдаем предупреждение
                    if text.Count == 0:
                        self.set_status("Технические требования пусты!")
                        messagebox.showwarning("Внимание", "Технические требования есть, но они пусты!")
                        return
                    
                    # Текстовое содержимое для вывода в редактор
                    formatted_text = self.parse_tech_req(text)
                    
                    # Обновляем текстовое поле
                    self.current_reqs_text.delete(1.0, tk.END)
                    self.current_reqs_text.insert(tk.END, formatted_text)
                    
                    doc_name = active_doc.Name
                    self.set_status(f"Технические требования загружены из {doc_name}")
                    return
                except Exception as e:
                    # Если не удалось получить через первый метод, пробуем альтернативный
                    self.set_status("Пробуем альтернативный метод получения тех. требований...")
                    print(f"Exception in method 1: {str(e)}")
                
                # Альтернативный метод через ActiveDocument2D
                kompas_doc2D = self.app7.ActiveDocument2D()
                if kompas_doc2D:
                    tech_req = kompas_doc2D.TechnicalDemand()
                    if tech_req:
                        # Получаем количество строк (если такой метод есть)
                        if hasattr(tech_req, 'Count'):
                            count = tech_req.Count()
                            
                            # Если технические требования пустые
                            if count == 0:
                                self.set_status("В документе отсутствуют технические требования!")
                                messagebox.showwarning("Внимание", "В документе отсутствуют технические требования!")
                                return
                            
                            # Получаем текст технических требований в чистом виде
                            text = ""
                            for i in range(count):
                                line_text = tech_req.GetLine(i)
                                text += line_text + "\n"
                        else:
                            # Для API, не имеющего метода Count
                            self.set_status("Используем нестандартный метод получения тех. требований")
                            try:
                                # Пробуем получить текст напрямую из свойства Text
                                text_obj = tech_req.Text
                                
                                text = ""
                                for i in range(text_obj.Count):
                                    line = text_obj.TextLines[i]
                                    text += line.Str + "\n"
                            except Exception as e2:
                                self.set_status(f"Не удалось получить текст требований: {str(e2)}")
                                messagebox.showerror("Ошибка", f"Не удалось получить текст требований: {str(e2)}")
                                return
                        
                        # Обновляем текстовое поле
                        self.current_reqs_text.delete(1.0, tk.END)
                        self.current_reqs_text.insert(tk.END, text)
                        
                        doc_name = active_doc.Name
                        self.set_status(f"Технические требования загружены из {doc_name} (метод 2)")
                        return
                
            except Exception as e:
                error_message = self.handle_kompas_error(e, "получения технических требований")
                self.set_status("Ошибка при получении тех. требований")
                messagebox.showerror("Ошибка", error_message)
                print(f"Exception details: {str(e)}")
                
        except Exception as e:
            error_message = self.handle_kompas_error(e, "работы с документом")
            self.set_status("Ошибка при работе с документом")
            messagebox.showerror("Ошибка", error_message)
            print(f"Exception details: {str(e)}")
            
    def save_technical_requirements(self):
        """Сохранение технических требований в активный документ"""
        # Используем общий метод с флагом сохранения документа
        self.apply_technical_requirements(save_document=True)
            
    def apply_technical_requirements(self, save_document=False):
        """Применение технических требований к активному документу без сохранения файла"""
        try:
            if not hasattr(self, 'module7') or not self.module7:
                self.connect_to_kompas()
                if not hasattr(self, 'module7') or not self.module7:
                    return
            
            # Получаем активный документ
            active_doc = self.app7.ActiveDocument
            if not active_doc:
                self.set_status("Нет активного документа")
                messagebox.showwarning("Внимание", "Нет активного документа в КОМПАС-3D")
                return
                
            # Получаем текст из редактора
            text_content = self.current_reqs_text.get(1.0, tk.END).strip()
            
            try:
                # Получаем интерфейс чертежа
                drawing_document = self.module7.IDrawingDocument(active_doc)
                
                # Получаем интерфейс технических требований
                tech_demand = drawing_document.TechnicalDemand
                
                # Если текст пустой, очищаем технические требования и выходим
                if not text_content:
                    # Если технические требования уже созданы, очищаем их
                    if tech_demand.IsCreated:
                        text_obj = tech_demand.Text
                        while text_obj.Count > 0:
                            # Находим первую строку и удаляем её
                            line = text_obj.TextLines[0]
                            line.Delete()
                        
                        # Применяем изменения
                        if hasattr(tech_demand, 'Update'):
                            tech_demand.Update()
                        
                        # Обновляем документ
                        if hasattr(drawing_document, 'Update'):
                            drawing_document.Update()
                        else:
                            # Если метода Update нет, пробуем обновить через active_doc
                            if hasattr(active_doc, 'Update'):
                                active_doc.Update()
                        
                        self.set_status("Технические требования очищены")
                        return
                    else:
                        # Если требования не созданы и текст пустой, просто выходим
                        self.set_status("Нет технических требований для применения")
                        return
                
                # Если технические требования не созданы, создаем их
                if not tech_demand.IsCreated:
                    tech_demand.Create()
                
                # Очищаем текущие технические требования
                text_obj = tech_demand.Text
                while text_obj.Count > 0:
                    # Находим первую строку и удаляем её
                    line = text_obj.TextLines[0]
                    line.Delete()
                
                # Разбиваем текст на строки
                lines = text_content.split("\n")
                
                # Удаляем пустые строки
                lines = [line.strip() for line in lines if line.strip()]
                
                # Удаляем существующую нумерацию и определяем, какие строки должны быть пронумерованы
                cleaned_lines = []
                should_number = []
                
                for i, line in enumerate(lines):
                    # Удаляем существующую нумерацию (если есть)
                    clean_line = re.sub(r'^\d+\.\s*', '', line)
                    # Удаляем отступы в начале строки
                    clean_line = clean_line.lstrip()
                    cleaned_lines.append(clean_line)
                    
                    # Определяем, должна ли строка иметь номер
                    # Строка не должна иметь номер, если она начинается с маленькой буквы или с тире/дефиса
                    # и не является первой строкой
                    if i > 0 and (
                        (len(clean_line) > 0 and clean_line[0].islower()) or 
                        clean_line.startswith('-') or 
                        clean_line.startswith('–')
                    ):
                        should_number.append(False)
                    else:
                        should_number.append(True)
                
                # Добавляем строки с информацией о нумерации
                for i, (line, should_num) in enumerate(zip(cleaned_lines, should_number)):
                    processed_lines.append((line, should_num))
            except Exception as e:
                self.handle_kompas_error(e, "обработки технических требований")
                return
            else:
                # Если автонумерация выключена, просто используем текст как есть
                for line in lines:
                    # Проверяем, есть ли нумерация в начале строки
                    num_match = re.match(r'^(\d+)\.\s*(.*)', line)
                    if num_match:
                        # Если есть нумерация, извлекаем текст и информацию о нумерации
                        req_text = num_match.group(2).strip()
                        processed_lines.append((req_text, True))
                    else:
                        # Проверяем, есть ли отступ в начале строки (для продолжения пункта)
                        indent_match = re.match(r'^\s+(.+)', line)
                        if indent_match:
                            # Если есть отступ, это продолжение пункта
                            req_text = indent_match.group(1).strip()
                            processed_lines.append((req_text, False))
                        else:
                            # Если нет нумерации и отступа, используем строку как есть
                            processed_lines.append((line, True))  # Предполагаем, что это новый пункт
                
            # Добавляем каждую строку в технические требования KOMPAS-3D
            for i, (line_text, is_numbered) in enumerate(processed_lines):
                try:
                    # Добавляем строку
                    text_line = text_obj.Add()
                    text_line.Str = line_text
                    
                    # Устанавливаем нумерацию
                    if is_numbered:
                        text_line.Numbering = 1
                    else:
                        text_line.Numbering = 0
                        
                except Exception as line_error:
                    print(f"Ошибка при добавлении строки '{line_text}': {str(line_error)}")
                    # Продолжаем с следующей строкой
                
            # Применяем изменения
            if hasattr(tech_demand, 'Update'):
                tech_demand.Update()
            
            # Обновляем документ
            if hasattr(drawing_document, 'Update'):
                drawing_document.Update()
            else:
                # Если метода Update нет, пробуем обновить через active_doc
                if hasattr(active_doc, 'Update'):
                    active_doc.Update()
            
            # Сохраняем документ, если нужно
            if save_document:
                try:
                    active_doc.Save()
                    self.set_status("Документ сохранен")
                except Exception as e:
                    error_msg = self.handle_kompas_error(e, "сохранения документа")
                    self.set_status("Не удалось сохранить документ автоматически")
            
            doc_name = active_doc.Name
            self.set_status(f"Технические требования применены к {doc_name}" + 
                          (" и сохранены" if save_document else " (без сохранения файла)"))
            
            if save_document:
                messagebox.showinfo("Информация", f"Технические требования успешно сохранены в {doc_name}")
            else:
                messagebox.showinfo("Информация", f"Технические требования успешно применены к {doc_name} (без сохранения файла)")
                
        except Exception as e:
            error_message = self.handle_kompas_error(e, "применения технических требований")
            self.set_status("Ошибка при применении тех. требований")
            messagebox.showerror("Ошибка", error_message)
            print(f"Exception details: {str(e)}")
            
    def select_document_in_tree(self, document):
        """Выбор документа в дереве документов"""
        try:
            if not document:
                return
                
            doc_name = document.Name
            
            # Ищем документ в дереве
            for item in self.doc_tree.get_children():
                if self.doc_tree.item(item, 'text') == doc_name:
                    # Выделяем найденный документ
                    self.doc_tree.selection_set(item)
                    self.doc_tree.see(item)
                    return
                    
            # Если документ не найден в дереве, обновляем дерево
            self.update_documents_tree()
            
            # Пробуем найти документ снова
            for item in self.doc_tree.get_children():
                if self.doc_tree.item(item, 'text') == doc_name:
                    self.doc_tree.selection_set(item)
                    self.doc_tree.see(item)
                    return
                    
        except Exception as e:
            # Игнорируем ошибки при выборе документа в дереве
            pass
            
    def update_documents_tree(self, search_term=None):
        """Обновление дерева документов"""
        try:
            if not hasattr(self, 'app7') or not self.app7:
                self.set_status("Нет подключения к KOMPAS-3D")
                return
                
            # Очистка дерева
            for item in self.doc_tree.get_children():
                self.doc_tree.delete(item)
                
            # Получение списка документов
            documents = self.app7.Documents
            doc_count = 0
            
            # Заполнение дерева
            for i in range(documents.Count):
                doc = documents.Item(i)
                doc_name = doc.Name
                
                # Если задан поисковый запрос, фильтруем документы
                if search_term and search_term.lower() not in doc_name.lower():
                    continue
                    
                # Определяем тип документа
                doc_type = "Неизвестный тип"
                try:
                    # Проверяем, является ли документ чертежом
                    doc2D_s = doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['IDrawingDocument'], 
                                                       pythoncom.IID_IDispatch)
                    doc_type = "Чертеж"
                except:
                    try:
                        # Проверяем, является ли документ 3D-моделью
                        doc3D_s = doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['IDocument3D'], 
                                                           pythoncom.IID_IDispatch)
                        doc_type = "3D-модель"
                    except:
                        try:
                            # Проверяем, является ли документ спецификацией
                            spec_s = doc._oleobj_.QueryInterface(self.module7.NamesToIIDMap['ISpecificationDocument'], 
                                                              pythoncom.IID_IDispatch)
                            doc_type = "Спецификация"
                        except:
                            pass
                
                # Получаем путь к документу
                doc_path = doc.Path
                if not doc_path:
                    doc_path = "Документ не сохранен"
                
                # Добавляем документ в дерево
                item_id = self.doc_tree.insert("", "end", text=doc_name, 
                                             values=(doc_type, doc_path))
                
                # Если это активный документ, выделяем его
                if self.app7.ActiveDocument and doc.Name == self.app7.ActiveDocument.Name:
                    self.doc_tree.selection_set(item_id)
                    self.doc_tree.see(item_id)
                
                doc_count += 1
                
            # Обновляем информацию о количестве документов
            self.set_status(f"Найдено документов: {doc_count}")
            self.docs_count_label.config(text=f"Документов: {doc_count}")
            
        except Exception as e:
            self.set_status(f"Ошибка при обновлении дерева документов: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось обновить дерево документов: {str(e)}")
            
    def periodic_update(self):
        """Периодическое обновление информации о документах"""
        try:
            # Проверяем, запущен ли KOMPAS-3D
            if self.is_kompas_running():
                # Обновляем информацию об активном документе
                self.update_active_document_info()
            else:
                # Если KOMPAS-3D не запущен, обновляем статус
                self.connect_status.config(text="🔴 Нет подключения", foreground='red')
                self.active_doc_label.config(text="Нет активного документа")
        except Exception as e:
            # Игнорируем ошибки при периодическом обновлении
            pass
            
        # Планируем следующее обновление
        self.root.after(5000, self.periodic_update)

    def format_text(self, format_type):
        """Форматирование выделенного текста"""
        try:
            current_selection = self.current_reqs_text.tag_ranges(tk.SEL)
            if not current_selection:
                self.set_status("Нет выделенного текста для форматирования")
                return
                
            start, end = current_selection
            
            # Проверяем, есть ли формат уже
            existing_tags = self.current_reqs_text.tag_names(start)
            
            if format_type in existing_tags:
                # Если формат уже есть - удаляем его
                self.current_reqs_text.tag_remove(format_type, start, end)
                self.set_status(f"Формат '{format_type}' удален")
            else:
                # Добавляем форматирование
                self.current_reqs_text.tag_add(format_type, start, end)
                self.set_status(f"Применен формат '{format_type}'")
                
        except Exception as e:
            self.set_status(f"Ошибка при форматировании текста: {str(e)}")
            
    def create_new_document(self, doc_type="drawing"):
        """Создание нового документа в KOMPAS-3D"""
        try:
            if not hasattr(self, 'app7') or not self.app7:
                self.connect_to_kompas()
                if not hasattr(self, 'app7') or not self.app7:
                    return
                    
            if doc_type == "drawing":
                # Создаем новый чертеж
                doc = self.app7.Document2D()
                doc.Create(False, True)  # Создаем новый документ без видимости и с новым окном
                doc_type_name = "чертеж"
            else:
                # Создаем новую 3D-модель
                doc = self.app7.Document3D()
                doc.Create(False, True)  # Создаем новый документ без видимости и с новым окном
                doc_type_name = "3D-модель"
            
            # Активируем документ
            doc.Active = True
            
            # Обновляем информацию
            self.update_active_document_info()
            self.update_documents_tree()
            
            self.set_status(f"Создан новый документ: {doc_type_name}")
            
        except Exception as e:
            error_message = self.handle_kompas_error(e, "создания нового документа")
            self.set_status(f"Ошибка при создании нового документа")
            messagebox.showerror("Ошибка", error_message)
            
    def show_new_document_dialog(self):
        """Показать диалог выбора типа нового документа"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Создание нового документа")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Центрирование окна
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Заголовок
        tk.Label(dialog, text="Выберите тип документа:", font=("Arial", 12)).pack(pady=10)
        
        # Кнопки
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        # Кнопка для создания чертежа
        drawing_btn = ttk.Button(button_frame, text="Чертеж", width=15, 
                              command=lambda: [dialog.destroy(), self.create_new_document("drawing")])
        drawing_btn.pack(side=tk.LEFT, padx=5)
        
        # Кнопка для создания 3D-модели
        model_btn = ttk.Button(button_frame, text="3D-модель", width=15, 
                            command=lambda: [dialog.destroy(), self.create_new_document("3d")])
        model_btn.pack(side=tk.LEFT, padx=5)
        
        # Кнопка отмены
        cancel_btn = ttk.Button(dialog, text="Отмена", width=15, 
                             command=dialog.destroy)
        cancel_btn.pack(pady=10)
        
    def disconnect_from_kompas(self):
        """Отключение от KOMPAS-3D"""
        try:
            if hasattr(self, 'app7') and self.app7:
                # Освобождаем COM-объекты
                self.app7 = None
                self.module7 = None
                self.api7 = None
                self.const7 = None
                
                # Вызываем сборщик мусора для освобождения COM-объектов
                gc.collect()
                
                # Обновляем статус
                self.connect_status.config(text="🔴 Нет подключения", foreground='red')
                self.set_status("Отключено от KOMPAS-3D")
                
                # Очищаем дерево документов
                for item in self.doc_tree.get_children():
                    self.doc_tree.delete(item)
                    
                return True
            else:
                self.set_status("Нет активного подключения к KOMPAS-3D")
                return False
                
        except Exception as e:
            self.set_status(f"Ошибка при отключении от KOMPAS-3D: {str(e)}")
            return False
            
    def on_closing(self):
        """Обработчик закрытия приложения"""
        try:
            # Отключаемся от KOMPAS-3D
            if hasattr(self, 'app7') and self.app7:
                self.disconnect_from_kompas()
                
            # Освобождаем COM
            pythoncom.CoUninitialize()
            
            # Закрываем приложение
            self.root.destroy()
            
        except Exception as e:
            print(f"Ошибка при закрытии приложения: {str(e)}")
            self.root.destroy()
            
    def handle_kompas_error(self, e, operation="операции"):
        """Обработка ошибок при работе с KOMPAS-3D"""
        error_msg = str(e)
        error_code = None
        
        # Извлекаем код ошибки из сообщения, если он есть
        if "0x" in error_msg:
            try:
                # Ищем шестнадцатеричный код ошибки
                match = re.search(r'0x[0-9A-Fa-f]+', error_msg)
                if match:
                    error_code = match.group(0)
            except:
                pass
                
        # Формируем сообщение об ошибке
        if error_code:
            message = f"Ошибка при выполнении {operation} в KOMPAS-3D.\n\nКод ошибки: {error_code}"
            
            # Добавляем описание для известных ошибок
            if error_code == "0x80004005":
                message += "\n\nНеуказанная ошибка. Возможно, проблема с доступом к объекту."
            elif error_code == "0x80020009":
                message += "\n\nИсключение в KOMPAS-3D. Проверьте состояние документа."
            elif error_code == "0x8002000A":
                message += "\n\nНеверный индекс или параметр."
            elif error_code == "0x80020006":
                message += "\n\nНеизвестное имя или метод."
            
            # Добавляем рекомендации
            message += "\n\nРекомендации:\n"
            message += "1. Убедитесь, что KOMPAS-3D запущен и работает корректно.\n"
            message += "2. Проверьте, что у вас есть права на редактирование документа.\n"
            message += "3. Попробуйте переподключиться к KOMPAS-3D."
            
        else:
            message = f"Ошибка при выполнении {operation} в KOMPAS-3D.\n\n{error_msg}"
            
        # Выводим сообщение в статусную строку
        self.set_status(f"Ошибка: {error_msg}")
        
        # Возвращаем сообщение для использования в диалогах
        return message
        
    def apply_list_formatting(self, tech_req):
        """Применение форматирования списком к техническим требованиям"""
        try:
            # Пробуем разные варианты форматирования в зависимости от версии API
            if hasattr(tech_req, 'FormatAsList'):
                tech_req.FormatAsList()
                # Установка параметров списка
                if hasattr(tech_req, 'ListParams'):
                    tech_req.ListParams = True
                # Установка типа списка (нумерованный)
                if hasattr(tech_req, 'ListType'):
                    tech_req.ListType = 0  # 0 - нумерованный список
                # Применяем нумерацию
                self.apply_numbering(tech_req)
                self.set_status("Применено форматирование списком и нумерация (метод 1)")
                return True
            elif hasattr(tech_req, 'Text') and hasattr(tech_req.Text, 'FormatAsList'):
                tech_req.Text.FormatAsList()
                # Установка параметров списка
                if hasattr(tech_req.Text, 'ListParams'):
                    tech_req.Text.ListParams = True
                # Установка типа списка (нумерованный)
                if hasattr(tech_req.Text, 'ListType'):
                    tech_req.Text.ListType = 0  # 0 - нумерованный список
                # Применяем нумерацию
                self.apply_numbering(tech_req)
                self.set_status("Применено форматирование списком и нумерация (метод 2)")
                return True
            else:
                # Если нет прямого метода, пробуем через интерфейс IText
                try:
                    text_obj = tech_req.Text
                    
                    # Устанавливаем параметры списка для всего текста
                    if hasattr(text_obj, 'ListParams'):
                        text_obj.ListParams = True
                    
                    # Установка типа списка для всего текста (нумерованный)
                    if hasattr(text_obj, 'ListType'):
                        text_obj.ListType = 0  # 0 - нумерованный список
                    
                    # Устанавливаем стиль списка для всех строк
                    for i in range(text_obj.Count):
                        line = text_obj.TextLines[i]
                        if hasattr(line, 'ListStyle'):
                            line.ListStyle = True
                        # Установка параметров списка для каждой строки
                        if hasattr(line, 'ListParams'):
                            line.ListParams = True
                        # Установка типа списка (нумерованный)
                        if hasattr(line, 'ListType'):
                            line.ListType = 0  # 0 - нумерованный список
                    
                    # Применяем нумерацию
                    self.apply_numbering(tech_req)
                    self.set_status("Применено форматирование списком и нумерация (метод 3)")
                    return True
                except Exception as e:
                    self.set_status(f"Не удалось применить форматирование списком: {str(e)}")
                    return False
                    
        except Exception as e:
            self.set_status(f"Ошибка при форматировании списком: {str(e)}")
            return False
            
    def apply_numbering(self, tech_req):
        """Применение нумерации к техническим требованиям средствами API KOMPAS"""
        try:
            # Пробуем разные варианты нумерации в зависимости от версии API
            if hasattr(tech_req, 'SetNumbering'):
                # Метод 1: Прямой вызов метода SetNumbering
                tech_req.SetNumbering()
                # Установка параметров списка
                if hasattr(tech_req, 'ListParams'):
                    tech_req.ListParams = True
                # Установка типа списка (нумерованный)
                if hasattr(tech_req, 'ListType'):
                    tech_req.ListType = 0  # 0 - нумерованный список
                # Включение автоматической нумерации
                if hasattr(tech_req, 'AutoNumbering'):
                    tech_req.AutoNumbering = True
                self.set_status("Применена нумерация (метод 1)")
                return True
            elif hasattr(tech_req, 'Text') and hasattr(tech_req.Text, 'SetNumbering'):
                # Метод 2: Вызов метода SetNumbering у объекта Text
                tech_req.Text.SetNumbering()
                # Установка параметров списка
                if hasattr(tech_req.Text, 'ListParams'):
                    tech_req.Text.ListParams = True
                # Установка типа списка (нумерованный)
                if hasattr(tech_req.Text, 'ListType'):
                    tech_req.Text.ListType = 0  # 0 - нумерованный список
                # Включение автоматической нумерации
                if hasattr(tech_req.Text, 'AutoNumbering'):
                    tech_req.Text.AutoNumbering = True
                self.set_status("Применена нумерация (метод 2)")
                return True
            elif hasattr(tech_req, 'Text') and hasattr(tech_req.Text, 'NumberingStyle'):
                # Метод 3: Установка свойства NumberingStyle
                tech_req.Text.NumberingStyle = True
                # Установка параметров списка
                if hasattr(tech_req.Text, 'ListParams'):
                    tech_req.Text.ListParams = True
                # Установка типа списка (нумерованный)
                if hasattr(tech_req.Text, 'ListType'):
                    tech_req.Text.ListType = 0  # 0 - нумерованный список
                # Включение автоматической нумерации
                if hasattr(tech_req.Text, 'AutoNumbering'):
                    tech_req.Text.AutoNumbering = True
                self.set_status("Применена нумерация (метод 3)")
                return True
            else:
                # Метод 4: Пробуем установить нумерацию для каждой строки
                try:
                    text_obj = tech_req.Text
                    
                    # Установка параметров списка для всего текста
                    if hasattr(text_obj, 'ListParams'):
                        text_obj.ListParams = True
                    
                    # Установка типа списка для всего текста (нумерованный)
                    if hasattr(text_obj, 'ListType'):
                        text_obj.ListType = 0  # 0 - нумерованный список
                    
                    # Включение автоматической нумерации для всего текста
                    if hasattr(text_obj, 'AutoNumbering'):
                        text_obj.AutoNumbering = True
                    
                    # Применяем нумерацию к каждой строке
                    for i in range(text_obj.Count):
                        line = text_obj.TextLines[i]
                        # Установка стиля нумерации
                        if hasattr(line, 'NumberingStyle'):
                            line.NumberingStyle = True
                        elif hasattr(line, 'Numbering'):
                            line.Numbering = True
                            
                        # Установка параметров списка для каждой строки
                        if hasattr(line, 'ListParams'):
                            line.ListParams = True
                            
                        # Установка типа списка для каждой строки (нумерованный)
                        if hasattr(line, 'ListType'):
                            line.ListType = 0  # 0 - нумерованный список
                    
                    # Дополнительная попытка установить параметры списка для всего текста
                    if hasattr(text_obj, 'ListParams'):
                        text_obj.ListParams = True
                        
                    self.set_status("Применена нумерация (метод 4)")
                    return True
                except Exception as e:
                    self.set_status(f"Не удалось применить нумерацию: {str(e)}")
                    return False
        except Exception as e:
            self.set_status(f"Ошибка при применении нумерации: {str(e)}")
            return False
            
    def apply_numbering(self):
        """Применение автоматической нумерации к техническим требованиям"""
        try:
            # Получаем текст из редактора
            text_content = self.current_reqs_text.get(1.0, tk.END).strip()
            
            if not text_content:
                messagebox.showinfo("Информация", "Нет текста для нумерации")
                return
                
            # Разбиваем текст на строки
            lines = text_content.split("\n")
            
            # Удаляем пустые строки
            lines = [line.strip() for line in lines if line.strip()]
            
            # Удаляем существующую нумерацию и определяем, какие строки должны быть пронумерованы
            cleaned_lines = []
            should_number = []
            
            for i, line in enumerate(lines):
                # Удаляем существующую нумерацию (если есть)
                clean_line = re.sub(r'^\d+\.\s*', '', line)
                cleaned_lines.append(clean_line)
                
                # Определяем, должна ли строка иметь номер
                # Строка не должна иметь номер, если она начинается с маленькой буквы или с тире/дефиса
                # и не является первой строкой
                if i > 0 and (
                    (len(clean_line) > 0 and clean_line[0].islower()) or 
                    clean_line.startswith('-') or 
                    clean_line.startswith('–')
                ):
                    should_number.append(False)
                else:
                    should_number.append(True)
            
            # Применяем новую нумерацию
            result_lines = []
            for i, (line, should_num) in enumerate(zip(cleaned_lines, should_number)):
                processed_lines.append((line, should_num))
            for i, (line, should_num) in enumerate(zip(cleaned_lines, should_number)):
                if should_num:
                    result_lines.append(f"{i+1}. {line}")
                else:
                    result_lines.append(line)
            
            # Обновляем текст в редакторе
            self.current_reqs_text.delete(1.0, tk.END)
            self.current_reqs_text.insert(1.0, "\n".join(result_lines))
            
            self.set_status("Автонумерация применена")
            
        except Exception as e:
            self.set_status(f"Ошибка при применении автонумерации: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось применить автонумерацию: {str(e)}")
            
    def toggle_auto_numbering(self, is_enabled=None):
        """Переключение режима автонумерации"""
        try:
            # Если параметр не передан, используем текущее значение переменной
            if is_enabled is None:
                is_enabled = self.auto_numbering_var.get()
            else:
                # Иначе устанавливаем переданное значение
                self.auto_numbering_var.set(is_enabled)
                
            # Если автонумерация включена, применяем её к тексту
            if is_enabled:
                # Применяем автонумерацию к тексту
                self.apply_auto_numbering()
                
                # Добавляем обработчик ввода
                self.current_reqs_text.bind("<Return>", self.handle_return_with_numbering)
            else:
                # Если автонумерация выключена, удаляем обработчик ввода
                self.current_reqs_text.unbind("<Return>")
                
        except Exception as e:
            self.set_status(f"Ошибка при переключении автонумерации: {str(e)}")
            print(f"Error toggling auto numbering: {str(e)}")
            
    def handle_return_with_numbering(self, event):
        """Обработка нажатия Enter при включенной автонумерации"""
        try:
            # Проверяем, включена ли автонумерация
            if not self.auto_numbering_var.get():
                return  # Если автонумерация выключена, используем стандартную обработку Enter
                
            # Получаем текущую позицию курсора
            cursor_pos = self.current_reqs_text.index(tk.INSERT)
            line, col = map(int, cursor_pos.split('.'))
            
            # Получаем текущую строку
            current_line = self.current_reqs_text.get(f"{line}.0", f"{line}.end").strip()
            
            # Получаем все строки текста до текущей позиции
            all_text_before = self.current_reqs_text.get(1.0, f"{line}.0").strip()
            lines_before = all_text_before.split("\n") if all_text_before else []
            
            # Подсчитываем количество пронумерованных строк до текущей
            numbered_lines_before = [l for l in lines_before if re.match(r'^\d+\.\s', l)]
            
            # Вставляем новую строку
            self.current_reqs_text.insert(tk.INSERT, "\n")
            
            # Определяем, нужно ли добавлять номер к новой строке
            # Если текущая строка начинается с номера, добавляем следующий номер
            if re.match(r'^\d+\.\s', current_line):
                # Извлекаем текущий номер
                current_num_match = re.match(r'^(\d+)\.', current_line)
                if current_num_match:
                    current_num = int(current_num_match.group(1))
                    next_number = current_num + 1
                else:
                    next_number = len(numbered_lines_before) + 1
                
                # Вставляем номер в новую строку
                self.current_reqs_text.insert(f"{line+1}.0", f"{next_number}. ")
                
                # Перемещаем курсор после номера
                self.current_reqs_text.mark_set(tk.INSERT, f"{line+1}.{len(str(next_number)) + 2}")
            
            # Предотвращаем стандартную обработку Enter
            return "break"
            
        except Exception as e:
            self.set_status(f"Ошибка при обработке ввода: {str(e)}")
            print(f"Error handling return with numbering: {str(e)}")
            
    def apply_auto_numbering(self):
        """Применение автоматической нумерации к техническим требованиям"""
        try:
            # Получаем текст из редактора
            text_content = self.current_reqs_text.get(1.0, tk.END).strip()
            
            if not text_content:
                return
                
            # Разбиваем текст на строки
            lines = text_content.split("\n")
            
            # Удаляем пустые строки
            lines = [line.strip() for line in lines if line.strip()]
            
            # Удаляем существующую нумерацию и определяем, какие строки должны быть пронумерованы
            cleaned_lines = []
            should_number = []
            
            for i, line in enumerate(lines):
                # Удаляем существующую нумерацию (если есть)
                clean_line = re.sub(r'^\d+\.\s*', '', line)
                cleaned_lines.append(clean_line)
                
                # Определяем, должна ли строка иметь номер
                # Строка не должна иметь номер, если она начинается с маленькой буквы или с тире/дефиса
                # и не является первой строкой
                if i > 0 and (
                    (len(clean_line) > 0 and clean_line[0].islower()) or 
                    clean_line.startswith('-') or 
                    clean_line.startswith('–')
                ):
                    should_number.append(False)
                else:
                    should_number.append(True)
            
            # Применяем новую нумерацию
            result_lines = []
            number_counter = 1
            
            for i, (line, should_num) in enumerate(zip(cleaned_lines, should_number)):
                if should_num:
                    result_lines.append(f"{number_counter}. {line}")
                    number_counter += 1
                else:
                    result_lines.append(f"    {line}")  # Добавляем отступ для ненумерованных строк
            
            # Обновляем текст в редакторе
            self.current_reqs_text.delete(1.0, tk.END)
            self.current_reqs_text.insert(1.0, "\n".join(result_lines))
            
            # Устанавливаем статус
            self.set_status("Автонумерация применена")
            
        except Exception as e:
            self.set_status(f"Ошибка при применении автонумерации: {str(e)}")
            print(f"Error applying auto numbering: {str(e)}")
            
    def remove_auto_numbering(self):
        """Удаление автоматической нумерации из технических требований"""
        try:
            # Получаем текст из редактора
            text_content = self.current_reqs_text.get(1.0, tk.END).strip()
            
            if not text_content:
                return
                
            # Разбиваем текст на строки
            lines = text_content.split("\n")
            
            # Удаляем нумерацию
            result_lines = []
            for line in lines:
                # Удаляем нумерацию в начале строки
                clean_line = re.sub(r'^\d+\.\s*', '', line)
                result_lines.append(clean_line)
            
            # Обновляем текст в редакторе
            self.current_reqs_text.delete(1.0, tk.END)
            self.current_reqs_text.insert(1.0, "\n".join(result_lines))
            
        except Exception as e:
            self.set_status(f"Ошибка при удалении автонумерации: {str(e)}")
            print(f"Error removing auto numbering: {str(e)}")
            
    def parse_tech_req(self, text_lines):
        """
        Парсинг технических требований из объекта TextLines в удобный формат
        :param text_lines: Объект TextLines из KOMPAS API
        :return: отформатированный текст технических требований с соблюдением нумерации
        """
        # Текстовое содержимое для вывода
        formatted_text = ""
        
        # Счетчик для нумерации требований
        count = 0
        
        # Текущее требование и его номер
        current_req = ""
        current_req_num = 0
        
        # Проходим по каждой строке технических требований
        i = 0
        while i < text_lines.Count:
            line = text_lines.TextLines[i]
            line_text = line.Str.strip()
            
            # Если строка пустая, переходим к следующей
            if not line_text:
                i += 1
                continue
            
            # Проверяем, есть ли нумерация у строки
            if line.Numbering == 1:
                # Если уже собрали предыдущее требование, добавляем его
                if current_req:
                    formatted_text += f"{current_req_num}. {current_req}\n"
                
                # Начинаем новое требование
                count += 1
                current_req_num = count
                current_req = line_text
            else:
                # Если текущая строка - продолжение предыдущей
                if current_req:
                    # Добавляем пробел перед продолжением, если это не начало строки и 
                    # предыдущая строка не заканчивается на знак переноса
                    if (not current_req.endswith(" ") and 
                        not current_req.endswith("-") and 
                        not line_text.startswith("-")):
                        current_req += " "
                    current_req += line_text
                else:
                    # Если это первая строка без нумерации, создаем новое требование
                    count += 1
                    current_req_num = count
                    current_req = line_text
            
            # Если это последняя строка, добавляем накопленное требование
            if i == text_lines.Count - 1 and current_req:
                formatted_text += f"{current_req_num}. {current_req}\n"
            
            i += 1
            
        return formatted_text
        
    def clean_tech_req_line(self, line):
        """Очистка строки технических требований от нумерации и форматирования"""
        # Удаляем нумерацию в начале строки (например, "1. ", "2. ", и т.д.)
        line = re.sub(r'^\s*\d+\.\s*', '', line)
        
        # Удаляем другие возможные маркеры списка
        line = re.sub(r'^\s*[•\-–—]\s*', '', line)
        
        # Удаляем лишние пробелы в начале и конце строки
        line = line.strip()
        
        return line
            
if __name__ == "__main__":
    # Инициализация COM
    pythoncom.CoInitialize()
    
    try:
        # Запуск приложения
        root = tk.Tk()
        app = KompasApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Критическая ошибка приложения: {str(e)}")
    finally:
        # Завершение COM
        pythoncom.CoUninitialize()
