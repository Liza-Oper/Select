import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
import copy

# Окно для работы с фильтрацией
# P.S. доделать логику И и ИЛИ (взять из старого кода)

class FilterWindow(tk.Toplevel):
    def __init__(self, parent, dataframe):
        super().__init__(parent)
        self.filtered_data = None
        self.dataframe = dataframe
        self.filters = []
        self.column = None
        self.logic_connection = tk.StringVar(value="AND")  # Default logic
        self.condition = tk.StringVar(value="Не выбрано")
        self.value_filter = []
        self.filtered_values = []
        self.search_var = tk.StringVar()
        self.date_calendar = None

        self.iconbitmap('icon.ico')
        self.title("Фильтрация")
        self.geometry("600x800")
        self.configure(bg="#ed7aa0")

        style = ttk.Style()
        style.configure("TButton", padding=8, relief="flat", background="#ed7aa0", foreground="white", font=("Arial", 10))
        style.configure("TCombobox", padding=5, font=("Arial", 10), relief="flat", background="#ffffff")
        style.configure("TEntry", padding=5, font=("Arial", 10), relief="flat", fieldbackground="#ed7aa0")
        style.configure("TListbox", font=("Arial", 10), relief="flat", background="#f9f9f9")
        style.configure("TLabel", font=("Arial", 10), background="#f4f4f4")

        # Dropdown for selecting a column
        tk.Label(self, text="Выберите колонку:", bg="#f8f1ff").pack(pady=5)
        self.column_dropdown = ttk.Combobox(self, values=list(dataframe.columns))
        self.column_dropdown.pack(pady=5)
        self.column_dropdown.bind("<<ComboboxSelected>>", self.load_values)

        # Dropdown for conditions
        tk.Label(self, text="Фильтровать по условию:", bg="#f8f1ff").pack(pady=5)
        self.condition_dropdown = ttk.Combobox(self, textvariable=self.condition, values=[
            "Не выбрано", "Содержит данные", "Не содержит данные", "Текст содержит",
            "Текст не содержит", "Текст начинается с", "Текст заканчивается на",
            "Текст в точности", "Дата до", "Дата после", "Дата равна",
            "Число больше", "Число меньше", "Число равно",
        ])
        self.condition_dropdown.pack(pady=5)

        # Entry for input value
        tk.Label(self, text="Введите значение:", bg="#f8f1ff").pack(pady=5)
        self.value_entry = tk.Entry(self)
        self.value_entry.pack(pady=5)

        # Calendar button
        self.calendar_button = tk.Button(self, text="Выбрать дату", command=self.show_calendar)
        self.calendar_button.pack(pady=5)

        # Logic selection (AND/OR)
        tk.Label(self, text="Выберите связь между фильтрами:", bg="#f8f1ff").pack(pady=5)
        self.logic_and_button = tk.Radiobutton(self, text="И", variable=self.logic_connection, value="AND", bg="#f8f1ff")
        self.logic_or_button = tk.Radiobutton(self, text="ИЛИ", variable=self.logic_connection, value="OR", bg="#f8f1ff")
        self.logic_and_button.pack(pady=5)
        self.logic_or_button.pack(pady=5)

        # Filter list display
        self.filter_listbox = tk.Listbox(self, height=10)
        self.filter_listbox.pack(pady=10, fill="both", expand=True)

        # Buttons
        tk.Button(self, text="Добавить фильтр", command=self.add_filter).pack(side="left", padx=5, pady=5)
        tk.Button(self, text="Применить", command=self.apply_filter).pack(side="left", padx=5, pady=5)
        tk.Button(self, text="Отмена", command=self.destroy).pack(side="right", padx=5, pady=5)

    def show_calendar(self):
        # Отображает календарь для выбора даты
        if not self.date_calendar:
            self.date_calendar = DateEntry(self, width=12, background="darkblue", foreground="white", borderwidth=2)
            self.date_calendar.place(x=self.calendar_button.winfo_x() + self.calendar_button.winfo_width() + 10,
                                      y=self.calendar_button.winfo_y())  # Размещаем справа от кнопки
            self.date_calendar.bind("<<DateEntrySelected>>", self.on_date_selected)

    def on_date_selected(self, event=None):
        # Обработчик выбора даты
        selected_date = self.date_calendar.get_date().strftime("%Y-%m-%d")
        self.value_entry.delete(0, tk.END)  # Очищаем текущее значение
        self.value_entry.insert(0, selected_date)  # Вставляем выбранную дату

    def load_values(self, event=None):
        # Загрузка уникальных значений для выбранной колонки
        self.column = self.column_dropdown.get()
        if self.column:
            unique_values = self.dataframe[self.column].dropna().unique()
            self.value_filter = list(map(str, unique_values))

    def add_filter(self):
        # Добавление фильтра
        column = self.column_dropdown.get()
        condition = self.condition_dropdown.get()
        value = self.value_entry.get()
        logic = self.logic_connection.get()

        if not column or condition == "Не выбрано" or not value:
            messagebox.showwarning("Ошибка", "Заполните все поля.")
            return

        self.filters.append({"column": column, "condition": condition, "value": value, "logic": logic})
        self.filter_listbox.insert(tk.END, f"{column} {condition} {value} ({logic})")

    def apply_filter(self):
        # Применение фильтров
        if not self.filters:
            messagebox.showwarning("Ошибка", "Добавьте хотя бы один фильтр перед применением.")
            return

        filtered_df = self.dataframe
        combined_condition = None

        for filter_ in self.filters:
            column = filter_["column"]
            condition = filter_["condition"]
            value = filter_["value"]
            logic = filter_["logic"]

            # Создание условия фильтрации
            if condition == "Текст содержит":
                condition_series = filtered_df[column].astype(str).str.contains(value, na=False)
            elif condition == "Текст не содержит":
                condition_series = ~filtered_df[column].astype(str).str.contains(value, na=False)
            elif condition == "Дата до":
                condition_series = pd.to_datetime(filtered_df[column]) < pd.to_datetime(value)
            elif condition == "Дата после":
                condition_series = pd.to_datetime(filtered_df[column]) > pd.to_datetime(value)
            elif condition == "Дата равна":
                condition_series = pd.to_datetime(filtered_df[column]) == pd.to_datetime(value)
            else:
                condition_series = filtered_df[column] == value  # По умолчанию фильтрация по значению

            # Объединение условий с логикой И/ИЛИ
            if combined_condition is None:
                combined_condition = condition_series
            else:
                if logic == "AND":
                    combined_condition &= condition_series
                elif logic == "OR":
                    combined_condition |= condition_series

        if combined_condition is not None:
            self.filtered_data = filtered_df[combined_condition]
            self.destroy()
        else:
            messagebox.showwarning("Ошибка", "Не удалось применить фильтры.")


# Класс для работы с сортировкой
# По возможности пересмотреть дизайн
# conditions_met = []
        # for column in self.columns:
        #     if selected_conditions == "Текст содержит":
        #         conditions_met.append(filtered_df[column].astype(str).str.contains(input_value, na=False))
        #     elif selected_conditions == "Текст не содержит":
        #         conditions_met.append(~filtered_df[column].astype(str).str.contains(input_value, na=False))
        #     elif selected_conditions == "Текст начинается с":
        #         conditions_met.append(filtered_df[column].astype(str).str.startswith(input_value, na=False))
        #     elif selected_conditions == "Текст заканчивается на":
        #         conditions_met.append(filtered_df[column].astype(str).str.endswith(input_value, na=False))
        #     elif selected_conditions == "Текст в точности":
        #         conditions_met.append(filtered_df[column].astype(str) == input_value)
        #     elif selected_conditions == "Дата до":
        #         conditions_met.append(pd.to_datetime(filtered_df[column]) < pd.to_datetime(input_value))
        #     elif selected_conditions == "Дата после":
        #         conditions_met.append(pd.to_datetime(filtered_df[column]) > pd.to_datetime(input_value))
        #     elif selected_conditions == "Дата равна":
        #         conditions_met.append(pd.to_datetime(filtered_df[column]) == pd.to_datetime(input_value))
        #     elif selected_conditions == "Число больше":
        #         conditions_met.append(filtered_df[column] > float(input_value))
        #     elif selected_conditions == "Число меньше":
        #         conditions_met.append(filtered_df[column] < float(input_value))
        #     elif selected_conditions == "Число равно":
        #         conditions_met.append(filtered_df[column] == float(input_value))
        #
        # # Применяем связь И/ИЛИ
        # if logic_operator == "И":
        #     filter_condition = conditions_met[0]
        #     for cond in conditions_met[1:]:
        #         filter_condition &= cond
        # else:  # Логическое ИЛИ
        #     filter_condition = conditions_met[0]
        #     for cond in conditions_met[1:]:
        #         filter_condition |= cond
        #
        # self.filtered_data = filtered_df[filter_condition]
        # self.destroy()

class SortWindow(tk.Toplevel):
    def __init__(self, parent, dataframe):
        super().__init__(parent)
        self.title("Сортировка данных")
        self.geometry("500x400")
        self.configure(bg="#ffb455")
        self.sorted_data = None
        self.dataframe = dataframe
        self.sort_columns = []
        self.iconbitmap('icon.ico')

        # Интерфейс
        tk.Label(self, text="Выберите колонку для сортировки:").pack(pady=5)
        self.column_dropdown = ttk.Combobox(self, values=list(dataframe.columns), state="readonly")
        self.column_dropdown.pack(pady=5)

        tk.Label(self, text="Порядок сортировки:").pack(pady=5)
        self.order_dropdown = ttk.Combobox(self, values=["По возрастанию", "По убыванию"], state="readonly")
        self.order_dropdown.pack(pady=5)

        tk.Button(self, text="Добавить поле", command=self.add_sort_field).pack(pady=10)
        tk.Button(self, text="Применить", command=self.apply_sort).pack(pady=5)
        tk.Button(self, text="Отмена", command=self.destroy).pack(pady=5)

        # Список выбранных полей
        self.sort_list = tk.Listbox(self, height=8)
        self.sort_list.pack(pady=10, fill="both", expand=True)
    # Добавление колонок
    def add_sort_field(self):
        column = self.column_dropdown.get()
        order = self.order_dropdown.get()

        if not column or not order:
            messagebox.showwarning("Ошибка", "Выберите колонку и порядок сортировки!")
            return

        self.sort_columns.append((column, order))
        self.sort_list.insert("end", f"{column} ({order})")
    # Приминить сортировку
    def apply_sort(self):
        if not self.sort_columns:
            messagebox.showwarning("Ошибка", "Добавьте хотя бы одно поле для сортировки!")
            return

        try:
            ascending_list = [order == "По возрастанию" for _, order in self.sort_columns]
            column_list = [column for column, _ in self.sort_columns]

            self.sorted_data = self.dataframe.sort_values(by=column_list, ascending=ascending_list)
            messagebox.showinfo("Успех", "Сортировка применена!")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось выполнить сортировку: {e}")
# КЛасс для работы с окно группировки
class GroupWindow(tk.Toplevel):
    def __init__(self, parent, dataframe):
        super().__init__(parent)
        self.title("Группировка данных")
        self.geometry("600x500")
        self.configure(bg="#1c8954")
        self.grouped_data = None
        self.dataframe = dataframe
        self.group_fields = []  # Список для хранения (поле, функция)
        self.iconbitmap('icon.ico')

        # Элементы интерфейса
        tk.Label(self, text="Выберите колонку для группировки:", bg="#1c8954", fg="white").pack(pady=5)
        self.column_dropdown = ttk.Combobox(self, values=list(dataframe.columns), state="readonly")
        self.column_dropdown.pack(pady=5)

        tk.Label(self, text="Выберите агрегирующую функцию:", bg="#1c8954", fg="white").pack(pady=5)
        self.function_dropdown = ttk.Combobox(self, values=["Сумма", "Минимум", "Максимум", "Среднее", "Количество"], state="readonly")
        self.function_dropdown.pack(pady=5)

        tk.Button(self, text="Добавить поле", command=self.add_group_field).pack(pady=10)

        tk.Label(self, text="Выбранные поля и функции:", bg="#1c8954", fg="white").pack(pady=5)
        self.group_list = tk.Listbox(self, height=10)
        self.group_list.pack(pady=10, fill="both", expand=True)

        tk.Button(self, text="Применить", command=self.apply_group).pack(pady=10)
        tk.Button(self, text="Отмена", command=self.destroy).pack(pady=5)

    # Добавление выбранного поля с функцией
    def add_group_field(self):
        column = self.column_dropdown.get()
        agg_function = self.function_dropdown.get()

        if not column:
            messagebox.showwarning("Ошибка", "Выберите колонку!")
            return

        if not agg_function:
            messagebox.showwarning("Ошибка", "Выберите агрегирующую функцию!")
            return

        # Добавляем комбинацию (поле, функция)
        self.group_fields.append((column, agg_function))
        self.group_list.insert("end", f"{column} ({agg_function})")

    # Применение группировки
    def apply_group(self):
        if not self.group_fields:
            messagebox.showwarning("Ошибка", "Добавьте хотя бы одно поле и функцию!")
            return

        # Привязка функций Pandas к пользовательскому выбору
        agg_mapping = {
            "Сумма": "sum",
            "Минимум": "min",
            "Максимум": "max",
            "Среднее": "mean",
            "Количество": "size"
        }

        # Формируем словарь для агрегирования
        agg_dict = {}
        for column, agg_function in self.group_fields:
            if agg_function == "Количество":
                # Для функции "Количество" отдельно обрабатываем
                if column not in agg_dict:
                    agg_dict[column] = []
                agg_dict[column].append("size")
            else:
                if column not in agg_dict:
                    agg_dict[column] = []
                agg_dict[column].append(agg_mapping[agg_function])

        try:
            # Группируем данные
            grouped = self.dataframe.groupby(list({col for col, _ in self.group_fields}))
            self.grouped_data = grouped.agg(agg_dict).reset_index()

            messagebox.showinfo("Успех", "Группировка применена!")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось выполнить группировку: {e}")
class SpreadsheetApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Таблица - Sheets ☺ ")
        self.geometry("800x600")
        self.configure(bg="#6bd1a0")
        self.iconbitmap('icon.ico') #иконка котика (есть в каждом классе Window)
        self.file_path = None
        self.dataframe = None
        self.history = []  # Для кнопок "Назад" и "Вперед"
        self.future = []   # Для отмененных действий


        # Верхняя панель с кнопками
        button_frame = tk.Frame(self, bg="#6bd1a0")
        button_frame.pack(side="top", fill="x", pady=10)

        # Стиль кнопок
        button_style = {
            "bg": "#ffb455",  # Цвет фона
            "fg": "white",  # Цвет текста
            "font": ("Arial", 12, "bold"),
            "relief": "flat",  # Плоский стиль кнопки
            "bd": 1,  # Граница
            "highlightthickness": 0,  # Убираем эффект выделения
            "activebackground": "#b97825",  # Цвет фона при наведении
            "activeforeground": "black",  # Цвет текста при наведении
            "padx": 20,  # Горизонтальные отступы
            "pady": 6,  # Вертикальные отступы
            "width": 12,  # Ширина кнопки
            "height": 1,  # Высота кнопки
            "highlightcolor": "#ffffff",  # Цвет при фокусе
            "borderwidth": 1,  # Толщина границы
            # "relief": "solid"  # Сильно выраженная граница
        }
        # Кнопки
        tk.Button(button_frame, text="Открыть файл", command=self.open_file, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Сохранить файл", command=self.save_file, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Фильтрация", command=self.open_filter_window, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Сортировка", command=self.open_sort_window, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Группировка", command=self.open_group_window, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Сбросить", command=self.reset_actions, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Назад", command=self.undo_action, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Вперед", command=self.redo_action, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Колонки", command=self.open_column_selection_window, **button_style).pack(
            side="left", padx=5)
        # кнопка для перехода в полноэкранный режим
        self.fullscreen_button = tk.Button(self, text="⬛", font=("Arial", 10), command=self.toggle_fullscreen)
        self.fullscreen_button.place(relx=1.0, rely=1.0, anchor="se", x=-50,
                                     y=-10)  # Размещение внизу справа с учетом отступов

        # self.fullscreen_button.bind("<Enter>", self.show_hint)
        self.fullscreen_button.bind("<Leave>", self.hide_hint)

        self.hint_label = tk.Label(self, text="Полноэкранный режим", bg="white", font=("Arial", 8), padx=5, pady=5)
        self.hint_label.place_forget()  # Скрываем подсказку по умолчанию

        style = ttk.Style()

        # Стили
        # Тривью
        style.configure("Custom.Treeview",
                        background="#FFFFFF",
                        foreground="#333333",
                        fieldbackground="#F4F4F4",
                        rowheight=30,
                        borderwidth=0)

        # Заголовки
        style.configure("Custom.Treeview.Heading",
                        background="#FFFFFF",
                        foreground="#4CAF50",
                        font=("Arial", 10, "bold"),
                        borderwidth=0)

        # Скроллбар
        style.configure("Custom.Vertical.TScrollbar",
                        gripcount=0,
                        background="#E0E0E0",
                        darkcolor="#E0E0E0",
                        lightcolor="#E0E0E0",
                        troughcolor="#F5F5F5")

        # Таблица
        self.tree = ttk.Treeview(self, show="headings", style="Custom.Treeview")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        self.scroll_y = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll_y.set)
        self.scroll_y.pack(side="right", fill="y")
        self.scroll_y.configure(style="Custom.Vertical.TScrollbar")

    def toggle_fullscreen(self): #переключение в фуллскрин

        is_fullscreen = self.attributes("-fullscreen")
        if is_fullscreen:
            self.attributes("-fullscreen", False)  # Выход из полноэкранного режима
            self.fullscreen_button.config(text="⬛")
        else:
            self.attributes("-fullscreen", True)  # Вкл полноэкранного режима
            self.fullscreen_button.config(text="🟩")
    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if not self.file_path:
            return

        if self.file_path.endswith(".xlsx"):
            self.dataframe = pd.read_excel(self.file_path)
        elif self.file_path.endswith(".csv"):
            self.dataframe = pd.read_csv(self.file_path)

        self.history = [self.dataframe.copy()]
        self.future.clear()
        self.display_table()

    def show_hint(self, event):  # Показываем подсказку
        self.hint_label.place(x=self.winfo_width() - 150, y=self.winfo_height() - 40)

    def hide_hint(self, event):  # Скрываем подсказку
        self.hint_label.place_forget()

    def save_file(self):
        if self.tree.get_children():  # Проверяем, есть ли данные в Treeview
            # Получение данных из Treeview
            data = []
            for item in self.tree.get_children():
                data.append(self.tree.item(item, "values"))

            # Получение заголовков столбцов
            columns = [self.tree.heading(col, "text") for col in self.tree["columns"]]

            # Создаем DataFrame
            dataframe = pd.DataFrame(data, columns=columns)

            # Сохранение файла
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
            if save_path:
                if save_path.endswith(".xlsx"):
                    dataframe.to_excel(save_path, index=False)
                elif save_path.endswith(".csv"):
                    dataframe.to_csv(save_path, index=False)

                messagebox.showinfo("Успех", "Файл сохранен!")
        else:
            messagebox.showwarning("Ошибка", "Нет данных для сохранения!")

    def open_column_selection_window(self):
        if self.dataframe is None:
            return

        column_selection_window = ColumnSelectionWindow(self, self.dataframe.columns)
        self.wait_window(column_selection_window)

        selected_columns = column_selection_window.get_selected_columns()
        if selected_columns:
            self.update_column_display(selected_columns)

    def update_column_display(self, selected_columns):
        self.tree["columns"] = selected_columns
        for col in selected_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in self.dataframe[selected_columns].iterrows():
            self.tree.insert("", "end", values=row.tolist())

    def display_table(self):
        if self.dataframe is None:
            return

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.dataframe.columns)

        for col in self.dataframe.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in self.dataframe.iterrows():
            self.tree.insert("", "end", values=row.tolist())

    def open_filter_window(self):
        if self.dataframe is not None:
            filter_window = FilterWindow(self, self.dataframe)
            self.wait_window(filter_window)
            if filter_window.filtered_data is not None:
                self.update_history(filter_window.filtered_data)

    def open_sort_window(self):
        if self.dataframe is not None:
            sort_window = SortWindow(self, self.dataframe)
            self.wait_window(sort_window)
            if sort_window.sorted_data is not None:
                self.update_history(sort_window.sorted_data)

    def open_group_window(self):
        if self.dataframe is not None:
            group_window = GroupWindow(self, self.dataframe)
            self.wait_window(group_window)
            if group_window.grouped_data is not None:
                self.update_history(group_window.grouped_data)
    # Обновление истории действий
    def update_history(self, new_data):
        self.history.append(self.dataframe.copy())
        self.dataframe = new_data
        self.future.clear()
        self.display_table()
    # Метод для кнопки "Назад"
    def undo_action(self):
        if len(self.history) > 1:
            self.future.append(self.history.pop())
            self.dataframe = self.history[-1].copy()
            self.display_table()
        else:
            messagebox.showwarning("Ошибка", "Нет действий для отмены!")
    # Метод для кнопки "Вперед"
    def redo_action(self):
        if self.future:
            self.history.append(self.future.pop())
            self.dataframe = self.history[-1].copy()
            self.display_table()
        else:
            messagebox.showwarning("Ошибка", "Нет действий для повтора!")

    def reset_actions(self):
        if self.history:
            self.dataframe = self.history[0].copy()
            self.history = [self.dataframe]
            self.future.clear()
            self.display_table()

# class ExcelViewer(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.title("Excel Viewer")
#         self.geometry("800x600")
#         self.configure(bg="#f8f1ff")  # Пастельно-розовый фон
#
#         self.filepath = None
#         self.dataframes = {}  # Словарь для хранения DataFrame для каждого листа
#         self.current_sheet = None
#
#         # Кнопка загрузки файла
#         load_button = tk.Button(self, text="Загрузить файл", command=self.load_file, bg="#d8bfd8", fg="black")
#         load_button.pack(pady=10)
#
#         # Выпадающий список для выбора листа
#         self.sheet_dropdown = ttk.Combobox(self, state="readonly")
#         self.sheet_dropdown.bind("<<ComboboxSelected>>", self.switch_sheet)
#         self.sheet_dropdown.pack(pady=5)
#
#         # Таблица для отображения данных
#         self.table_frame = tk.Frame(self)
#         self.table_frame.pack(fill="both", expand=True)
#         self.table = ttk.Treeview(self.table_frame)
#         self.table.pack(fill="both", expand=True, side="left")
#         scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
#         scrollbar.pack(side="right", fill="y")
#         self.table.configure(yscrollcommand=scrollbar.set)
#
#     def load_file(self):
#         """Загружает Excel файл и читает все листы."""
#         self.filepath = filedialog.askopenfilename(
#             filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
#         )
#         if not self.filepath:
#             return
#
#         try:
#             # Проверка формата файла
#             if self.filepath.endswith(".xlsx"):
#                 self.dataframes = pd.read_excel(self.filepath, sheet_name=None)  # Чтение всех листов
#             elif self.filepath.endswith(".csv"):
#                 df = pd.read_csv(self.filepath)
#                 self.dataframes = {"Sheet1": df}  # Если это CSV, создаем один лист
#
#             # Настраиваем выпадающий список с листами
#             self.sheet_dropdown["values"] = list(self.dataframes.keys())
#             self.sheet_dropdown.current(0)  # Выбираем первый лист
#             self.switch_sheet()  # Отображаем первый лист
#
#         except Exception as e:
#             messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {e}")
#
#     def switch_sheet(self, event=None):
#         """Переключение между листами и обновление таблицы."""
#         selected_sheet = self.sheet_dropdown.get()
#         if selected_sheet in self.dataframes:
#             self.current_sheet = selected_sheet
#             self.display_table(self.dataframes[selected_sheet])
#
#     def display_table(self, dataframe):
#         """Отображает данные в таблице."""
#         # Очистка текущей таблицы
#         for col in self.table["columns"]:
#             self.table.delete(col)
#         self.table.delete(*self.table.get_children())
#
#         # Настройка колонок
#         self.table["columns"] = list(dataframe.columns)
#         self.table["show"] = "headings"  # Отображение только заголовков колонок
#
#         for col in dataframe.columns:
#             self.table.heading(col, text=col)
#             self.table.column(col, width=100, anchor="center")
#
#         # Добавление данных
#         for _, row in dataframe.iterrows():
#             self.table.insert("", "end", values=list(row))
class ColumnSelectionWindow(tk.Toplevel):
    def __init__(self, parent, columns):
        super().__init__(parent)
        self.title("Выбор колонок")
        self.geometry("400x800")
        self.configure(bg="#6bd1a0")
        self.iconbitmap('icon.ico')
        self.selected_columns = []
        button_style = {
            "bg": "#e64bc8",  # Цвет фона
            "fg": "white",  # Цвет текста
            "font": ("Arial", 12, "bold"),
            "relief": "flat",  # Плоский стиль кнопки
            "bd": 1,  # Граница
            "highlightthickness": 0,  # Убираем эффект выделения
            "activebackground": "#85066D",  # Цвет фона при наведении
            "activeforeground": "black",  # Цвет текста при наведении
            "padx": 20,  # Горизонтальные отступы
            "pady": 6,  # Вертикальные отступы
            "width": 12,  # Ширина кнопки
            "height": 1,  # Высота кнопки
            "highlightcolor": "#ffffff",  # Цвет при фокусе
            "borderwidth": 1,  # Толщина границы
            # "relief": "solid"  # Сильно выраженная граница
        }
        self.check_buttons = {}
        for col in columns:
            var = tk.BooleanVar(value=True)
            chk_btn = tk.Checkbutton(self, text=col, variable=var, bg="#6bd1a0", fg="black", font = "bold" )
            chk_btn.pack(anchor="w", padx=10, pady=5)
            self.check_buttons[col] = var

        button_frame = tk.Frame(self, bg="#6bd1a0")
        button_frame.pack(side="bottom", pady=10)

        tk.Button(button_frame, text="Применить", command=self.apply_selection, **button_style).pack(side="left", padx=5)
        tk.Button(button_frame, text="Отменить", command=self.destroy, **button_style).pack(side="left", padx=5)

    def apply_selection(self):
        self.selected_columns = [col for col, var in self.check_buttons.items() if var.get()]
        self.destroy()

    def get_selected_columns(self):
        return self.selected_columns
def main():
    app = SpreadsheetApp()
    app.mainloop()


if __name__ == "__main__":
    main()