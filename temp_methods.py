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
