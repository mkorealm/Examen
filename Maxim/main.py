import os
import tkinter as tk
from tkinter import font
from tkinter import ttk
from tkinter.messagebox import showerror

bg_col = "#009990"  # Цвет фона
fg_col = "#990990"  # Цвет текств

script_dir = os.path.dirname(__file__)  # Возвращает текущую директорию


class Windows(tk.Tk):
    """Этот класс принимает в себя конфигурацию всех окон и имеет функцию переключения между фреймами"""

    def __init__(self: tk.Tk, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        # Center win
        window_width = self.winfo_reqwidth()
        window_height = self.winfo_reqheight()
        position_right = int(self.winfo_screenwidth() / 2.25 - window_width / 2)
        position_down = int(self.winfo_screenheight() / 3 - window_height / 2)

        self.wm_title("Королев Максим Андреевич")  # Устанавливает название окна

        # Родительский контейнер для всех окон
        container = tk.Frame(self, padx=200, pady=100)
        container.pack()

        self.frames = {}
        for F in (First, Second):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(column=0, row=0, sticky="nsew")

        self.show_frame(First)  # Устанавливает активный фрейм

    def show_frame(self, content):
        """Функция переключения между фреймами(классами)"""
        frame = self.frames[content]
        frame.tkraise()


class First(tk.Frame):
    """Этот класс выводит первый фрейм. Здесь выполняется функционал ввода текста в файл Excel"""

    def __init__(self: tk.Frame, parent, controller):
        tk.Frame.__init__(self, parent)

        # Импорт модуля в котором содержится функция споздания и заполнения Excel документа
        from Maxim.Modules.insert_zapis import insert_zapis

        default_font = font.Font(font="TkDefaultFont:")  # Стилизация текста

        # Родительский контейнер для поля ввода данных
        insert_container = tk.LabelFrame(self, borderwidth=0)
        insert_container.pack()

        # Добавление информационных строк для пользователя
        date_lbl = tk.Label(insert_container, text=f"Укажите дату", fg=fg_col, font=default_font)
        date_lbl.grid(column=0, row=0)
        FIO_lbl = tk.Label(insert_container, text=f"Введите ФИО", fg=fg_col, font=default_font)
        FIO_lbl.grid(column=0, row=1)
        aud_lbl = tk.Label(insert_container, text=f"Укажите фудиторию", fg=fg_col, font=default_font)
        aud_lbl.grid(column=0, row=2)
        pick_lbl = tk.Label(insert_container, text=f"Введите время получения", fg=fg_col, font=default_font)
        pick_lbl.grid(column=0, row=3)
        drop_lbl = tk.Label(insert_container, text=f"Введите время сдачи", fg=fg_col, font=default_font)
        drop_lbl.grid(column=0, row=4)

        # Добавление строк ввода для пользователя
        date_entry = tk.Entry(insert_container)
        date_entry.grid(column=1, row=0)
        FIO_entry = tk.Entry(insert_container)
        FIO_entry.grid(column=1, row=1)
        aud_entry = tk.Entry(insert_container)
        aud_entry.grid(column=1, row=2)
        pick_entry = tk.Entry(insert_container)
        pick_entry.grid(column=1, row=3)
        drop_entry = tk.Entry(insert_container)
        drop_entry.grid(column=1, row=4)

        def get_values() -> list:
            """Функция получения аргументов из вводимых полей"""
            lst = [date_entry.get(), FIO_entry.get(), aud_entry.get(), pick_entry.get(), drop_entry.get()]

            return lst

        # Кнопка отправки данных введённые пользователем
        btn_insert = tk.Button(self, text="Заполнить Excel", width=23, font=default_font,
                               command=lambda: insert_zapis(get_values()))
        btn_insert.pack()

        # Кнопка перехода на следующий фрейм
        btn = tk.Button(self, text="Перейти в следующий класс", width=23, font=default_font,
                        command=lambda: controller.show_frame(Second))
        btn.pack()


class Second(tk.Frame):
    """Этот класс выводит второй фрейм. Он выполняет функционал вывода данных из таблицы"""

    def __init__(self: tk.Frame, parent, controller):
        tk.Frame.__init__(self, parent)

        def select_table(file):
            """Функция вывода данных из Excel файла в таблицу tkinter"""
            import openpyxl
            columns = []
            values = []
            # Попытка открытия Excel файла и подраздела. Внутри цикл для вытягивания значений полей документа
            try:
                wb = openpyxl.load_workbook(filename=file)
                column = wb.worksheets[0]
                for row in column:
                    values.append(row[0].value)
                    columns.append(row[1].value)

                # Создание таблицы
                Treeview = ttk.Treeview(self, columns=columns, show="headings")

                Treeview.heading("Дата", text="Дата")
                Treeview.heading("ФИО", text="ФИО")
                Treeview.heading("Аудитория", text="Аудитория")
                Treeview.heading("Время получения", text="Время получения")
                Treeview.heading("Время сдачи", text="Время сдачи")

                Treeview.insert("", tk.END, values=values[0:5])
                Treeview.pack(pady=10)
            except Exception as ex:
                showerror("Ошибка!", str(ex))  # Вывод ошибки

        def select_one_row(file):
            """Функция вывода записи по полю 'дата' из Excel файла в таблицу tkinter"""
            import openpyxl
            columns = []
            values = []
            # Попытка открытия Excel файла и подраздела. Внутри цикл для вытягивания значений полей документа
            try:
                wb = openpyxl.load_workbook(filename=file)
                column = wb.worksheets[0]
                for row in column:
                    values.append(row[0].value)
                    columns.append(row[1].value)

                # Создание таблицы
                Treeview = ttk.Treeview(self, columns=columns[0], show="headings")

                Treeview.heading("Дата", text="Дата")

                Treeview.insert("", tk.END, values=values[0])
                Treeview.pack(pady=10)
            except Exception as ex:
                showerror("Ошибка!", str(ex))  # Вывод ошибки

        default_font = font.Font(font="TkDefaultFont:")  # Стилизация текста

        # Добавление кнопок для вызова функций
        select_one_row_btn = tk.Button(self, text="Отпечатать записи по полю 'дата'", width=27, font=default_font,
                                       command=lambda: select_one_row(script_dir + "\\Данные.xlsx"))
        select_one_row_btn.pack()
        select_btn = tk.Button(self, text="Отпечатать Excel файл", width=27, font=default_font,
                               command=lambda: select_table(script_dir + "\\Данные.xlsx"))
        select_btn.pack()
