# 🛠️ Редактор технических требований KOMPAS-3D  
![banner](icons/icon.png)  
*Интуитивный интерфейс для работы с техническими требованиями в KOMPAS-3D*  

---

## 🌟 Особенности  

| **Функция**               | **Описание**                                                                 |
|---------------------------|-----------------------------------------------------------------------------|
| 🖥️ **Интеграция с KOMPAS** | Автоподключение к KOMPAS-3D через COM-интерфейс.                            |
| 📄 **PDF-экспорт**         | Пакетное сохранение чертежей в PDF с автоматическим созданием подпапок.     |
| 🧩 **Шаблоны**             | Готовые текстовые блоки с фильтрацией, категориями и вариантами ввода.      |
| 🔄 **Автообновление**      | Динамическое обновление списка документов каждую секунду.                   |
| 📋 **Автонумерация**       | Умная нумерация пунктов с поддержкой вложенности.                           |
| ✅ **Проверка ТТ**         | Проверка последовательности ТТ текущего документа и всех чертежей с индикацией статуса. |
| 🌙 **Темы оформления**     | Светлая и тёмная темы для комфортной работы.                                |
| 📂 **Управление файлами**  | Контекстное меню для работы с документами: открытие папки, закрытие и др.  |
| ⌨️ **Горячие клавиши**     | Настраиваемые горячие клавиши для всех действий через окно настроек.        |

---

## 🚀 Быстрый старт  

### 📋 Требования  
- **ОС**: Windows 10/11  
- **ПО**: KOMPAS-3D (API 7+)  
- **Зависимости**:  
  ```bash  
  pip install PyQt6 pywin32  
  ```  

### ⚙️ Установка  
1. 📥 Скачайте скрипт `kompas_app.py`.  
2. 🛠️ Установите зависимости:  
   ```bash  
   pip install -r requirements.txt  
   ```  
3. ▶️ Запустите:  
   ```bash  
   python main.py  
   ```  

---

## 🎮 Интерфейс  
![UI Preview](icons/GUI.png)  
*Главное окно приложения с панелями документов, редактором и шаблонами*  

---

## 📌 Ключевые функции  

### 📂 Работа с документами  
- **Дерево документов**:  
  - Новый столбец **"Статус"** с индикаторами:  
    - 🟢 — ТТ корректны.  
    - 🟡 — Обнаружены ошибки в ТТ.  
    - ⚪ — Документ не проверен или не является чертежом.  
  - Колонки: "Статус", "Имя", "Тип", "Путь".  
- **Поиск документов** 🔍 — фильтрация по имени в реальном времени.  
- **Контекстное меню** 🖱️:  
  - Открытие папки с документом.  
  - Закрытие документа с сохранением изменений.  
- **Двойной клик**: Активация документа, для чертежей — загрузка ТТ.  
- **PDF-экспорт** 🖨️:  
  - Одиночный файл: `Ctrl+Shift+S`.  
  - Пакетный экспорт: кнопка 📚 на панели.  

### ✨ Шаблоны  
- **Категории и варианты**:  
  ```python  
  {  
    "Общие": [  
      {"text": "Материал: сталь", "variants": ["нержавеющая", "углеродистая"]},  
      {"text": "Термообработка...", "variants": []}  
    ],  
    "Покрытия": [  
      {"text": "Цинкование", "variants": [{"text": "толщина {} мкм", "custom_input": True}]},  
      {"text": "Анодирование", "variants": []}  
    ]  
  }  
  ```  
- **Редактирование**: Открыть JSON через `Инструменты -> Редактировать шаблоны`.  

### ✅ Проверка последовательности ТТ  
- **Проверка текущего документа**:  
  - Кнопка ✅ на панели инструментов проверяет последовательность ТТ активного чертежа.  
  - Если есть нарушения:  
    - Показывается окно с описанием ошибок и правильной последовательностью.  
    - Кнопка "ОК" закрывает окно.  
    - Кнопка "Копировать" копирует правильный порядок в буфер обмена.  
  - Если ошибок нет:  
    - Сообщение "Последовательность и формат ТТ корректны" отображается зеленым в статус-баре.  
- **Проверка всех чертежей**:  
  - Кнопка 🛠️ на панели проверяет ТТ всех открытых чертежей.  
  - Результаты отображаются в столбце "Статус" дерева документов:  
    - 🟢 — ТТ корректны.  
    - 🟡 — Есть ошибки (всплывающая подсказка с деталями).  
  - При наличии ошибок открывается окно с описанием проблем по каждому документу.  

### 🛠️ Форматирование  
| **Кнопка** | **Действие**       | **Горячая клавиша** |  
|------------|--------------------|---------------------|  
| **B**      | Жирный текст       | `Ctrl+B`            |  
| *I*        | Курсив             | `Ctrl+I`            |  
| _U_        | Подчеркивание      | `Ctrl+U`            |  

### ⌨️ Настройка горячих клавиш  
- Откройте окно настроек через `Файл -> Настройки` (по умолчанию `Ctrl+I`).  
- Перейдите на вкладку **"Горячие клавиши"**.  
- Дважды щелкните на значении горячей клавиши, чтобы изменить её.  
- Примеры допустимых комбинаций: `Ctrl+S`, `Alt+F4`, `F5`.  
- Настройки сохраняются в `~/KOMPAS-TR/settings.json`.  

---

## 🚨 Troubleshooting  

### Распространенные ошибки  
| **Проблема**               | **Решение**                                   |  
|----------------------------|-----------------------------------------------|  
| Нет подключения к KOMPAS   | Запустите KOMPAS-3D от имени администратора.  |  
| Ошибки COM-интерфейса      | Переподключитесь через `Ctrl+K`.              |  
| Шаблоны не загружаются     | Удалите `templates.json` для пересоздания.    |  
| Документ не сохраняется    | Проверьте права на запись в папку.            |  
| Файл настроек устарел      | Удалите `settings.json`, он будет пересоздан. |  
| Статус в дереве некорректен| Нажмите 🛠️ для проверки всех чертежей.        |  

---

## 📦 Сборка в EXE  
```bash  
pyinstaller --onefile --windowed --icon=app.ico kompas_app.py  
```  
- 🎯 **Иконка**: Добавьте файл `app.ico` для кастомизации.  
- 📂 Готовый EXE: в папке `dist/`.  

---

## 📬 Контакты  
**Поддержка**:  
[![Telegram](https://img.shields.io/badge/Telegram-%40HeagBoKaT-blue)](https://t.me/HeagBoKaT)  
**Версия**: `1.2.0` | **Год**: 2025  

---

## 📜 Полная документация  

### 🔄 Периодическое обновление  
- Список документов обновляется каждую 1 сек.  
- Для ручного обновления нажмите `F6`.  

### ⚙️ Конфигурация  
- Файл шаблонов:  
  ```bash  
  ~/KOMPAS-TR/templates.json  
  ```  
- Настройки темы и горячих клавиш:  
  ```bash  
  ~/KOMPAS-TR/settings.json  
  ```  
- **Обновление настроек**:  
  Если файл `settings.json` не соответствует текущей версии разметки (например, отсутствуют горячие клавиши), он автоматически перезаписывается с настройками по умолчанию.  

### 🆕 Последние изменения  
- **Столбец "Статус"**: Добавлен в дерево документов для индикации состояния ТТ (🟢, 🟡, ⚪).  
- **Проверка всех чертежей**: Новая кнопка 🛠️ на панели инструментов для анализа ТТ всех открытых чертежей с обновлением индикаторов в дереве.  
