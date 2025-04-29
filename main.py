import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import subprocess

from tkcalendar import DateEntry
from datetime import datetime
import pandas as pd

from general_vars import general_vars, general_vars_dict

from functions import add_new_db as func_add_new_db
from functions import work_with_db as func_work_with_db
from functions import work_with_result_db as func_work_with_result_db

class NewPromedApp:
    def __init__(self, root):
        self.root = root
        self.root.state('zoomed')
        self.root.resizable(False, False)
        self.root.title("Перепись БУЗ ВО ЧДГП №1")
        self.root.iconbitmap('illustrations/nurse_icon.ico')
        self.current_frame = None

        # стили кнопок
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure(".", 
            background="white", foreground="black", fieldbackground="white"
        )
        self.style.configure("Main.TButton", font=("Helvetica", 16),
            foreground="black", background="white",
            padding=10, width=50, height=3
        )
        self.style.configure("MainGreen.TButton", font=("Helvetica", 16),
            background="#BAEBBA", bordercolor="#008500", foreground="black",
            padding=10, width=50, height=3
        )
        self.style.configure("Back.TButton", font=("Helvetica", 10),
            foreground="black", background="white",
            padding=5, width=25, height=3
        )
        self.style.configure("TFrame", background="white")
        self.style.configure("TLabel", background="white")
        self.style.configure("Green.TButton",
            background="#BAEBBA", bordercolor="#008500", foreground="black",
        )
        self.style.configure("Red.TButton",
            background="#FFB6C1", bordercolor="#FF0000", foreground="black"
        )
        self.style.configure("TEntry", fieldbackground="white")
        self.style.configure("TCombobox",
            fieldbackground="white", background="white", foreground="black"
        )
        self.style.map("TCombobox",
            fieldbackground=[("readonly", "white")],
            background=[("readonly", "white")]
        )
        self.style.configure("Treeview",
            background="white", fieldbackground="white"
        )
        self.style.map("Treeview",
            background=[("selected", "#0078d7")]
        )
        self.style.configure("Horizontal.TScrollbar",
            background="white", troughcolor="white", bordercolor="white", arrowcolor="black"
        )
        self.style.configure("Vertical.TScrollbar",
            background="white", troughcolor="white", bordercolor="white", arrowcolor="black"
        )
        self.style.configure("Protected.TEntry",
            fieldbackground="#f0f0f0", foreground="#666666", bordercolor="#cccccc", insertbackground="#666666"
        )


        # фон
        self.original_bg_image = Image.open("illustrations/bg_1.png")
        self.bg_label = tk.Label(root)
        self.bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)

        # изменения
        self.root.bind("<Configure>", self.on_window_resize)

        # экраны
        self.screen_main = tk.Frame(self.root)
        self.screen_add_new_db = tk.Frame(self.root)
        self.screen_work_with_db = tk.Frame(self.root)
        self.screen_work_with_result_db = tk.Frame(self.root)

        # словарь для хранения фонов
        self.bg_labels = {
            "screen_main": tk.Label(self.screen_main),
            "screen_add_new_db": tk.Label(self.screen_add_new_db),
            "screen_work_with_db": tk.Label(self.screen_work_with_db),
            "screen_work_with_result_db": tk.Label(self.screen_work_with_result_db)
        }

        # Добавление виджетов на экраны
        self.add_widgets_main()
        self.add_widgets_add_new_db()

        self.tree = None
        self.df = None
        self.search_entry = None
        self.column_var = None
        self.result_count = None
        self.add_widgets_work_with_db()

        self.report_var = None
        self.add_widgets_work_with_result_db()

        # Показ главного экрана
        self.show_screen(self.screen_main, 'screen_main')

    def add_widgets_main(self):
        # screen_main
        screen_main_center_frame = tk.Frame(self.screen_main, background="white")
        screen_main_center_frame.place(relx=0.5, rely=0.5, anchor="center")
        if self.check_have_db():
            ttk.Button(
                screen_main_center_frame, 
                text="Создать новую перепись",
                style="Main.TButton",
                command=lambda: self.show_screen(self.screen_add_new_db, 'screen_add_new_db')
            ).grid(row=0, column=0, padx=10, pady=10)
            ttk.Button(
                screen_main_center_frame, 
                text="Открыть перепись", 
                style="MainGreen.TButton",
                command=lambda: self.work_with_df()
            ).grid(row=1, column=0, padx=10, pady=10)
            ttk.Button(
                screen_main_center_frame, 
                text="Открыть отчёты", 
                style="MainGreen.TButton",
                command=lambda: self.work_with_result_df(from_main_screen=True)
            ).grid(row=2, column=0, padx=10, pady=10)

        else:
            ttk.Button(
                screen_main_center_frame, 
                text="Создать новую перепись",
                style="MainGreen.TButton",
                command=lambda: self.show_screen(self.screen_add_new_db, 'screen_add_new_db')
            ).grid(row=0, column=0, padx=10, pady=10)

        # фон главного экрана
        self.update_bg_image()

    def add_widgets_add_new_db(self):
        # screen_add_new_db
        ttk.Button(
            self.screen_add_new_db, 
            text="Вернуться в главное меню",
            style="Back.TButton", 
            command=lambda: self.show_screen(self.screen_main, 'screen_main')
        ).place(relx=0.1, rely=0.9, anchor="center")

        screen_center_frame = tk.Frame(self.screen_add_new_db, background="white")
        screen_center_frame.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Button(
            screen_center_frame, 
            text="Загрузить таблицу",
            style="Main.TButton", 
            command=lambda: self.add_new_db()
        ).grid(row=0, column=0, padx=10, pady=10)
        ttk.Button(
            screen_center_frame, 
            text="Загрузить таблицу из Промед",
            style="Main.TButton", 
            command=lambda: self.add_new_db_promed()
        ).grid(row=1, column=0, padx=10, pady=10)
        ttk.Button(
            screen_center_frame, 
            text="Создать новую пустую таблицу",
            style="Main.TButton", 
            command=lambda: self.add_new_empty_db()
        ).grid(row=2, column=0, padx=10, pady=10)

        # фон главного экрана
        self.update_bg_image()

    
    def add_widgets_work_with_db(self):
        # screen_work_with_db
        self.screen_work_with_db.configure(background="white")
        
        control_frame = ttk.Frame(
            self.screen_work_with_db,
            style="TFrame"
        )
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            control_frame,
            text="Добавить запись",
            command=self.add_new_record
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            control_frame,
            text="Сохранить изменения",
            command=func_work_with_db.save_changes,
            style="Green.TButton"
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            control_frame, 
            text="Отменить изменения", 
            command=self.cancel_changes
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            control_frame, 
            text="Вернуться в главное меню",
            command=lambda: self.show_screen(self.screen_main, 'screen_main')
        ).pack(side=tk.RIGHT, padx=5)

        search_frame = ttk.Frame(
            self.screen_work_with_db, 
            style="TFrame"
        )
        search_frame.pack(pady=10, fill=tk.X, padx=10)

        ttk.Label(
            search_frame, 
            text="Поиск:", 
            style="TLabel"
        ).pack(side=tk.LEFT, padx=5)

        self.search_entry = ttk.Entry(search_frame, 
                                width=50, 
                                style="TEntry"
        )
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind('<Return>', self.search_data)

        self.column_var = tk.StringVar(value="Все столбцы")
        columns = ["Все столбцы"] + list(general_vars.FILE_COLUMNS_LIST.value)
        ttk.Combobox(
            search_frame, 
            textvariable=self.column_var, 
            values=columns, 
            state="readonly", 
            style="TCombobox"
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            search_frame, 
            text="Найти", 
            command=self.search_data
        ).pack(side=tk.LEFT, padx=5)

        tree_frame = ttk.Frame(self.screen_work_with_db, style="TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree_vscroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_vscroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree_hscroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_hscroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            yscrollcommand=tree_vscroll.set,
            xscrollcommand=tree_hscroll.set,
            columns=general_vars.FILE_COLUMNS_LIST.value,
            show="headings",
            selectmode='browse',
            style="Treeview"
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_vscroll.config(command=self.tree.yview)
        tree_hscroll.config(command=self.tree.xview)

        self.tree.bind('<Double-1>', self.edit_record)

        status_frame = ttk.Frame(self.screen_work_with_db, style="TFrame")
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        self.result_count = tk.StringVar(value="Найдено записей: 0")
        ttk.Label(
            status_frame, 
            textvariable=self.result_count,
            style="TLabel",
            font=('Arial', 10, 'bold')
        ).pack(side=tk.LEFT)

        self.df = func_work_with_db.open_db()
        self.search_data()

        # фон главного экрана
        self.update_bg_image()

    def delete_record(self, record_index, window):
        if messagebox.askyesno(
            "Подтверждение", 
            "Вы уверены, что хотите удалить эту запись?",
            icon='warning'
        ):
            self.df = self.df.drop(record_index).reset_index(drop=True)
            self.search_data()
            window.destroy()
            messagebox.showinfo("Успех", "Запись удалена")

    def add_widgets_work_with_result_db(self):
        # screen_work_with_result_db
        report_frame = ttk.Frame(
            self.screen_work_with_result_db, 
            style="TFrame"
        )
        report_frame.pack(pady=10, fill=tk.X, padx=10)

        self.report_var = tk.StringVar(value=general_vars.REPORT_NAMES.value[0])
        columns = list(general_vars.REPORT_NAMES.value)
        ttk.Combobox(
            report_frame, 
            textvariable=self.report_var, 
            values=columns, 
            state="readonly", 
            style="TCombobox"
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            report_frame, 
            text="Открыть", 
            command=self.open_report
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            report_frame, 
            text="Сохранить файл с отчетом",
            style="Green.TButton", 
            command=lambda: self.save_reports_file()
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            report_frame, 
            text="Вернуться в главное меню",
            style="TButton", 
            command=lambda: self.show_screen(self.screen_main, 'screen_main')
        ).pack(side=tk.RIGHT, padx=5)

        tree_frame = ttk.Frame(self.screen_work_with_result_db, style="TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree_vscroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree_vscroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree_hscroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_hscroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            yscrollcommand=tree_vscroll.set,
            xscrollcommand=tree_hscroll.set,
            # columns=general_vars.FILE_COLUMNS_LIST.value,
            show="headings",
            selectmode='browse',
            style="Treeview"
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        tree_vscroll.config(command=self.tree.yview)
        tree_hscroll.config(command=self.tree.xview)

        func_work_with_result_db.make_reports()
        self.open_report()

        # фон главного экрана
        self.update_bg_image()
        
    def check_have_db(self):
        file_path = general_vars.FULL_FILE_PATH.value
        if os.path.exists(file_path):
            print(f"{file_path} существует.")
            return True
        else:
            print(f"{file_path} не существует.")
            return False

    def show_screen(self, frame, current_screen_name):
        if self.current_frame is not None:
            self.current_frame.pack_forget()

        for screen_name, bg_label in self.bg_labels.items():
            bg_label.place_forget()
        
        if current_screen_name == 'screen_main':
            self.add_widgets_main()

        self.current_frame = frame
        self.bg_label = self.bg_labels[current_screen_name]
        self.bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.update_bg_image()    
        self.current_frame.pack(fill="both", expand=True)

    def on_window_resize(self, event):
        if event.widget == self.root:
            self.update_bg_image()

    def update_bg_image(self):
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        resized_image = self.original_bg_image.resize(
            (window_width, window_height),
            Image.Resampling.LANCZOS
        )

        self.bg_photo = ImageTk.PhotoImage(resized_image)
        self.bg_label.config(image=self.bg_photo)
        self.bg_label.image = self.bg_photo

    def add_new_db(self):
        file_path = filedialog.askopenfilename(
            initialdir="/",
            title="Выберите файл",
            filetypes=[("Таблицы", "*.xlsx")]
        )
        if file_path:
            print(f"Выбран файл: {file_path}")
            func_add_new_db.create_new_db(file_path = file_path, full_path = general_vars.FULL_FILE_PATH.value)
            print('Обработка файла завершена')
            self.work_with_df()

    def add_new_db_promed(self):
        file_path = filedialog.askopenfilename(
            initialdir="/",
            title="Выберите файл",
            filetypes=[("Таблицы формата Промед", "*.ods")]
        )
        if file_path:
            print(f"Выбран файл: {file_path}")
            func_add_new_db.create_new_promed_db(file_path = file_path, full_path = general_vars.FULL_FILE_PATH.value)
            print('Обработка файла завершена')
            self.work_with_df()

    def add_new_empty_db(self):
        func_add_new_db.create_new_empty_db(file_path = general_vars.FULL_FILE_PATH.value)
        print('Создание файла завершено')
        self.work_with_df()

    def work_with_df(self):
        self.df = func_work_with_db.open_db()
        self.show_screen(self.screen_work_with_db, 'screen_work_with_db')
        
    def work_with_result_df(self, from_main_screen=False):
        if from_main_screen:
            func_work_with_result_db.make_reports()

        func_work_with_result_db.open_db(general_vars.REPORT_NAMES.value[0])
        self.show_screen(self.screen_work_with_result_db, 'screen_work_with_result_db')
        
    def add_new_record(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавление новой записи")
        add_window.geometry("550x700")
        
        entries = {}
        for col in general_vars.FILE_COLUMNS_LIST.value:
            frame = tk.Frame(add_window, background="white")
            frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(frame, text=col, width=15).pack(side=tk.LEFT)
            
            if col in general_vars.PROTECTED_COLUMNS.value:
                entry = ttk.Entry(frame, style="Protected.TEntry")
                entry.insert(0, '')
                entry.config(state='readonly')
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry

            elif col == 'Комментарии':
                entry = ttk.Entry(frame)
                entry.insert(0, '')
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry           
                
            elif col in ['ДР', 'Прибыл', 'Выбыл']:
                date_entry = DateEntry(
                    frame, 
                    date_pattern='dd.mm.yyyy',
                    year=datetime.now().year,
                    month=datetime.now().month,
                    day=datetime.now().day
                )
                date_entry._set_text('')
                date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = date_entry

            elif col == 'Улица':
                street_var = tk.StringVar(value='')
                street_cb = ttk.Combobox(frame, textvariable=street_var, values=general_vars_dict['STREET_OPTIONS'])
                street_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = street_var

            elif col == 'Пол':
                gender_var = tk.StringVar(value='') 
                gender_cb = ttk.Combobox(frame, textvariable=gender_var, values=general_vars.GENDER_OPTIONS.value, state="readonly")
                gender_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = gender_var

            elif col == 'Орг-ть':
                org_var = tk.StringVar(value='') 
                org_cb = ttk.Combobox(frame, textvariable=org_var, values=general_vars.ORG_OPTIONS.value, state="readonly")
                org_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = org_var
            else:
                default_value = '' 
                entry = ttk.Entry(frame)
                entry.insert(0, default_value)
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry
        
        button_frame = tk.Frame(add_window, background="white")
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Button(
            button_frame, 
            text="Добавить", 
            style="Green.TButton",
            command=lambda: self.save_record(entries, add_window)
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Отмена", 
            style="ButtonDB.TButton", 
            command=add_window.destroy
        ).pack(side=tk.RIGHT)

    def edit_record(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        
        record_index = int(selected_item)
        edit_window = tk.Toplevel(self.screen_work_with_db)
        edit_window.title("Редактирование записи")
        edit_window.geometry("550x700")
        
        entries = {}
        for col in general_vars.FILE_COLUMNS_LIST.value:
            frame = tk.Frame(edit_window, background="white")
            frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Label(frame, text=col, width=15).pack(side=tk.LEFT)
            
            value = self.df.at[record_index, col]
            if pd.isna(value):
                value = ''
            
            if col in general_vars.PROTECTED_COLUMNS.value:
                entry = ttk.Entry(frame, style="Protected.TEntry")
                entry.insert(0, str(value))
                entry.config(state='readonly')
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry
            elif col == 'Комментарии':
                entry = ttk.Entry(frame)
                entry.insert(0, str(value))
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry
            elif col in ['Прибыл', 'Выбыл']:
                date_entry = DateEntry(
                    frame, 
                    date_pattern='dd.mm.yyyy',
                    year=datetime.now().year,
                    month=datetime.now().month,
                    day=datetime.now().day
                )
                date_entry._set_text('')
                if value and str(value).strip():
                    date_obj = datetime.strptime(str(value), '%Y/%m')
                    date_entry.set_date(date_obj)
                date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = date_entry
            elif col == 'ДР':
                try:
                    date_obj = datetime.strptime(value, '%d.%m.%Y')
                except:
                    date_obj = datetime.now()
                
                date_entry = DateEntry(frame, date_pattern='dd.mm.yyyy')
                date_entry.set_date(date_obj)
                date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = date_entry
            elif col == 'Улица':
                street_var = tk.StringVar(value=value)
                street_cb = ttk.Combobox(frame, textvariable=street_var, values=general_vars_dict['STREET_OPTIONS'])
                street_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = street_var
            elif col == 'Пол':
                gender_var = tk.StringVar(value=value)
                gender_cb = ttk.Combobox(frame, textvariable=gender_var, values=general_vars.GENDER_OPTIONS.value, state="readonly")
                gender_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = gender_var
            elif col == 'Орг-ть':
                org_var = tk.StringVar(value=value)
                org_cb = ttk.Combobox(frame, textvariable=org_var, values=general_vars.ORG_OPTIONS.value, state="readonly")
                org_cb.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = org_var
            else:
                entry = ttk.Entry(frame)
                entry.insert(0, str(value))
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                entries[col] = entry
        
        button_frame = tk.Frame(edit_window, background="white")
        button_frame.pack(fill=tk.X, padx=5, pady=10)

        ttk.Button(
            button_frame, 
            text="Удалить запись", 
            style="Red.TButton",
            command=lambda: self.delete_record(record_index, edit_window)
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Сохранить", 
            style="Green.TButton",
            command=lambda: self.save_record(entries, edit_window, record_index)
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Отмена", 
            style="ButtonDB.TButton", 
            command=edit_window.destroy
        ).pack(side=tk.RIGHT)

    def save_record(self, entries, window, record_index = None):
        try:
            new_data = func_work_with_db.save_new_record(entries)
            if record_index:
                for col in general_vars.FILE_COLUMNS_LIST.value:
                    self.df.at[record_index, col] = new_data[col]  
            else:
                self.df = pd.concat([self.df, pd.DataFrame([new_data])], ignore_index=True)
            self.search_data()
            window.destroy()
            messagebox.showinfo("Успех", "Изменения сохранены")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить изменения: {e}")


    def search_data(self, event=None):
        search_term = self.search_entry.get().strip()
        column = self.column_var.get()
        
        if not search_term:
            results = self.df
        else:
            if column == "Все столбцы":
                mask = self.df.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
            else:
                mask = self.df[column].astype(str).str.contains(search_term, case=False)
            results = self.df[mask]
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not hasattr(self, 'columns_configured'):
            for col in general_vars.FILE_COLUMNS_LIST.value:
                self.tree.heading(col, text=col)
            self.columns_configured = True
        
        if not results.empty:
            for _, row in results.iterrows():
                self.tree.insert("", "end", values=list(row), iid=str(_))
            
            font = tk.font.nametofont("TkDefaultFont")
            for col in general_vars.FILE_COLUMNS_LIST.value:
                header_width = font.measure(col)
                content_width = max([font.measure(str(val)) for val in results[col].astype(str)] or [0])
                total_width = max(header_width, content_width) + 20
                self.tree.column(col, width=min(total_width, 300), anchor=tk.W)
                
            self.result_count.set(f"Найдено записей: {len(results)}")
        else:
            self.result_count.set("Найдено записей: 0")

    def cancel_changes(self):
        if messagebox.askyesno("Подтверждение", "Отменить все изменения?"):
            self.df = func_work_with_db.open_db()
            self.search_data()
            messagebox.showinfo("Информация", "Изменения отменены")

    def open_report(self):
        report_term = self.report_var.get()
        if not report_term:
            report_term = general_vars.REPORT_NAMES.value[0]
        
        self.report_df = func_work_with_result_db.open_db(report_term)
        self.tree.configure(columns=self.report_df.columns)
        results = self.report_df
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not hasattr(self, 'columns_configured'):
            for col_idx, col in enumerate(self.report_df.columns):
                self.tree.heading(col_idx, text=col)
            self.columns_configured = True
        else:
            for col_idx, col in enumerate(self.report_df.columns):
                self.tree.heading(col_idx, text=col)
            self.columns_configured = True
        
        if not results.empty:
            for _, row in results.iterrows():
                self.tree.insert("", "end", values=list(row), iid=str(_))
            
            font = tk.font.nametofont("TkDefaultFont")
            for col_ind, col in enumerate(self.report_df.columns):
                header_width = font.measure(col)
                content_width = max([font.measure(str(val)) for val in results[col].astype(str)] or [0])
                total_width = max(header_width, content_width) + 20
                self.tree.column(col_ind, width=min(total_width, 200), anchor=tk.W)

    def save_reports_file(self, filepath=general_vars.REPORTS_FILE_PATH.value):
        if os.path.exists(filepath):
            subprocess.Popen(f'explorer /select,"{os.path.normpath(filepath)}"')
        else:
            messagebox.showerror("Ошибка", "Файл не найден")



root = tk.Tk()
app = NewPromedApp(root)
root.mainloop()