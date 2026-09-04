import os
import sys
from pathlib import Path
import tkinter as tk

from tkinter import ttk, messagebox

import pandas as pd

from openpyxl import Workbook

from openpyxl.styles import (
    Font,
    Alignment,
    Border,
    PatternFill
)

from openpyxl.styles.borders import Side as BorderSide

from openpyxl.utils import get_column_letter

# ==========================================
# ПУТЬ К ПАПКЕ ПРОГРАММЫ
# ==========================================

def get_program_folder():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    
    return Path(__file__).resolve().parent


# ==========================================
# ОТКРЫТЬ ПАПКУ В WINDOWS
# ==========================================

def open_folder(path):
    try:
        os.startfile(path)
    except Exception as error:
        messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{error}")


# ==========================================
# ДОСТАТЬ ИМЯ ЯЧЕЙКИ ИЗ ИМЕНИ ФАЙЛА
# Пример:
# SSR_TC_20260810_PXCn108-e18_02_LSV_C14.mpt
# -> PXCn108-e18
# ==========================================

def extract_cell_name(file_path):
    parts = file_path.stem.split("_")
    if len(parts) >= 4:
        return parts[3]
    return file_path.stem

#==========================================
#Чтение МТР файла
#==========================================

def read_mtp(file_path):

    header_lines = None

    #читаем начало файла
    with open(file_path,"r",encoding="cp1252", errors= "replace") as file:
        for line in file:
            if line.startswith("Nb header lines"):
                header_lines = int(
                line.split(":")[1].strip()
                )
                break

    #Если строка не найдена
    if header_lines is None:
        raise ValueError("Не найдено значение Nb header lines")

    #Читаем саму таблицу
    df = pd.read_csv(file_path,sep="/t", skiprows = header_lines - 1, encoding= "cp1252")

    #Удаление случайного лишнего столбца

    df= df.loc[:,~df.columns.astype(str).str.starswitch("unnamed")]

    return df



#==========================================
#Поиск необходимого стобца
#==========================================

def find_column(df,possible_names):
    for name in possible_names:
        if name in df.columns:
            return name

    raise ValueError("Не найен столбец" + ",".join(possible_names))



#=========================================
#Преобразование времиени в стобцах
#=========================================

def convert_time(time_series,unit):
    time = time_series.astype(float).copy()
    #Первое всегда будет 0

    if unit == "5":
        return time

    elif unit == "min":
        return time /60
    elif unit == "h":
        return time/ 3600

    return time



#===========================================
#Преобразователь напряжения
#===========================================

def convert_voltage(voltage_series, unit):
    voltage = voltage_series.astype(float).copy()
    if unit == "V":
        return voltage
    elif unit == "mW":
        return voltage * 1000
    return voltage



#===========================================
#Преобразование тока
#===========================================

def convert_current(current_series, unit , mass_mg = None):
    current = current_series.astype(float).copy()

    if unit == "mA":
        return current
    elif unit == "A":
        return current / 1000
    elif unit == "mA/g":
        if mass_mg is None or mass_mg <=0:
            raise("Необходимо указать массу!")
        mass_g = mass_mg / 1000
        return current / mass_g
    elif unit  == "A/g":
        if mass_mg is None or mass_mg <=0:
            raise ValueError("Укажите массу!")

        mass_g = mass_mg / 1000
        current_a = current/ 1000
        return current_a/ mass_g
    return current



#==========================================
#Получение необходимого типа файлов
#==========================================

def prepeare_experiment_data(df,experiment, x_unit, y_unit, mass_mg = None):

    if experiment == "LSV":
        ewe_column = find_column(df,["Ewe/V", "<Ewe/V]>"])
        currentr_column = find_column(df,["<I>/mA", "I/ mA"])

        x = convert_voltage(df[ewe_column], x_unit)
        y = convert_current(df[currentr_column], y_unit, mass_mg)

    elif experiment == "Cstv":
        time_column = find_column(df, ["time/s"])
        currentr_column = find_column(df ["<I>/mA","I/mA"])

        x = convert_time(df[time_column],x_unit)
        y = convert_current(df[currentr_column],y_unit,mass_mg)

    elif experiment == "GCPL":

        time_column = find_column(df,["time/s"])
        ewe_column = find_column(df,["Ewe/V", "<Ewe/V>"])

        x = convert_time(df[time_column],x_unit)
        y = convert_voltage(df[ewe_column],y_unit)

    elif experiment == "OCV":

        time_column = find_column(df,["time/s"])
        ewe_column = find_column(df,["Ewe/V", "<Ewe/V>"])

        x = convert_time(df[time_column],x_unit)
        y = convert_voltage(df[ewe_column],y_unit)

    else:

        raise ValueError(
            "Неизвестный тип эксперимента"
        )

    return x, y  

# ==========================================
# ЧТЕНИЕ .MPT
# ==========================================

def read_mpt(file_path):

    header_lines = None

    with open(
        file_path,
        "r",
        encoding="cp1252",
        errors="replace"
    ) as file:

        for line in file:

            if line.startswith("Nb header lines"):

                header_lines = int(
                    line.split(":", 1)[1].strip()
                )

                break

    if header_lines is None:
        raise ValueError(
            f"Не найдено Nb header lines в файле:\n{file_path.name}"
        )

    df = pd.read_csv(
        file_path,
        sep="\t",
        skiprows=header_lines - 1,
        encoding="cp1252"
    )

    # Удаляем пустые столбцы
    df = df.loc[
        :,
        ~df.columns.astype(str).str.startswith("Unnamed")
    ]

    return df


# ==========================================
# ПОИСК СТОЛБЦА
# ==========================================

def find_column(df, possible_names):

    for name in possible_names:

        if name in df.columns:
            return name

    raise ValueError(
        "Не найден нужный столбец.\n\n"
        "Искали:\n"
        + ", ".join(possible_names)
        + "\n\n"
        "В файле есть:\n"
        + ", ".join(df.columns.astype(str))
    )


# ==========================================
# ВРЕМЯ
# ==========================================

def convert_time(time_series, unit):

    time = time_series.astype(float).copy()

    # Нормируем начало времени к нулю
    time = time - time.iloc[0]

    if unit == "s":
        return time

    elif unit == "min":
        return time / 60

    elif unit == "h":
        return time / 3600

    return time


# ==========================================
# НАПРЯЖЕНИЕ
# ==========================================

def convert_voltage(voltage_series, unit):

    voltage = voltage_series.astype(float).copy()

    if unit == "V":
        return voltage

    elif unit == "mV":
        return voltage * 1000

    return voltage


# ==========================================
# ТОК
# ==========================================

def convert_current(current_series, unit, mass_mg=None):

    current = current_series.astype(float).copy()

    # Исходные данные прибора: mA

    if unit == "mA":
        return current

    elif unit == "A":
        return current / 1000

    elif unit == "mA/g":

        if mass_mg is None or mass_mg <= 0:
            raise ValueError(
                "Для mA/g необходимо указать массу."
            )

        mass_g = mass_mg / 1000

        return current / mass_g

    elif unit == "A/g":

        if mass_mg is None or mass_mg <= 0:
            raise ValueError(
                "Для A/g необходимо указать массу."
            )

        mass_g = mass_mg / 1000
        current_a = current / 1000

        return current_a / mass_g

    return current


# ==========================================
# ВЫБОР ДАННЫХ ПО ТИПУ ЭКСПЕРИМЕНТА
# ==========================================

def prepare_experiment_data(
    df,
    experiment,
    x_unit,
    y_unit,
    mass_mg=None
):

    # --------------------------------------
    # LSV
    # Ewe -> I
    # --------------------------------------

    if experiment == "LSV":

        ewe_column = find_column(
            df,
            [
                "Ewe/V",
                "<Ewe/V>"
            ]
        )

        current_column = find_column(
            df,
            [
                "<I>/mA",
                "I/mA"
            ]
        )

        x = convert_voltage(
            df[ewe_column],
            x_unit
        )

        y = convert_current(
            df[current_column],
            y_unit,
            mass_mg
        )


    # --------------------------------------
    # CstV
    # time -> I
    # --------------------------------------

    elif experiment == "CstV":

        time_column = find_column(
            df,
            ["time/s"]
        )

        current_column = find_column(
            df,
            [
                "<I>/mA",
                "I/mA"
            ]
        )

        x = convert_time(
            df[time_column],
            x_unit
        )

        y = convert_current(
            df[current_column],
            y_unit,
            mass_mg
        )


    # --------------------------------------
    # GCPL
    # time -> Ewe
    # --------------------------------------

    elif experiment == "GCPL":

        time_column = find_column(
            df,
            ["time/s"]
        )

        ewe_column = find_column(
            df,
            [
                "Ewe/V",
                "<Ewe/V>"
            ]
        )

        x = convert_time(
            df[time_column],
            x_unit
        )

        y = convert_voltage(
            df[ewe_column],
            y_unit
        )


    # --------------------------------------
    # OCV
    # time -> Ewe
    # --------------------------------------

    elif experiment == "OCV":

        time_column = find_column(
            df,
            ["time/s"]
        )

        ewe_column = find_column(
            df,
            [
                "Ewe/V",
                "<Ewe/V>"
            ]
        )

        x = convert_time(
            df[time_column],
            x_unit
        )

        y = convert_voltage(
            df[ewe_column],
            y_unit
        )

    else:

        raise ValueError(
            f"Неизвестный эксперимент: {experiment}"
        )

    return x, y



def create_excel(
    processed_data,
    experiment,
    x_name,
    y_name,
    x_unit,
    y_unit,
    results_folder
):

    # ==========================================
    # СОЗДАЁМ EXCEL
    # ==========================================

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = experiment

    column = 1


    # ==========================================
    # ЗАПИСЫВАЕМ ВСЕ ОБРАЗЦЫ
    # ==========================================

    for sample in processed_data:

        cell_name = sample["cell_name"]
        x = sample["x"]
        y = sample["y"]


        # --------------------------------------
        # Строка 1 — величины
        # --------------------------------------

        sheet.cell(
            row=1,
            column=column,
            value=x_name
        )

        sheet.cell(
            row=1,
            column=column + 1,
            value=y_name
        )


        # --------------------------------------
        # Строка 2 — единицы
        # --------------------------------------

        sheet.cell(
            row=2,
            column=column,
            value=x_unit
        )

        sheet.cell(
            row=2,
            column=column + 1,
            value=y_unit
        )


        # --------------------------------------
        # Строка 3 — название ячейки
        # --------------------------------------

        sheet.cell(
            row=3,
            column=column,
            value=cell_name
        )

        sheet.cell(
            row=3,
            column=column + 1,
            value=cell_name
        )


        # --------------------------------------
        # ДАННЫЕ
        # --------------------------------------

        max_length = max(
            len(x),
            len(y)
        )

        for index in range(max_length):

            excel_row = index + 4

            if index < len(x):

                sheet.cell(
                    row=excel_row,
                    column=column,
                    value=float(x.iloc[index])
                )

            if index < len(y):

                sheet.cell(
                    row=excel_row,
                    column=column + 1,
                    value=float(y.iloc[index])
                )


        # Следующий файл начинается
        # через два столбца
        column += 2


    # ==========================================
    # ПОЛНОЕ ОФОРМЛЕНИЕ EXCEL
    # ==========================================

    # --------------------------
    # ЦВЕТА
    # --------------------------

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    alternate_fill = PatternFill(
        fill_type="solid",
        fgColor="E7E6E6"
    )


    # --------------------------
    # ЛИНИИ
    # --------------------------

    thin_line = BorderSide(
        style="thin",
        color="FFA6A6A6"
    )

    thick_line = BorderSide(
        style="thick",
        color="FF000000"
    )


    # --------------------------
    # ШРИФТ
    # --------------------------

    header_font = Font(
        bold=True
    )


    # --------------------------
    # ОБХОДИМ ВСЮ ТАБЛИЦУ
    # --------------------------

    for row in range(
        1,
        sheet.max_row + 1
    ):

        for col in range(
            1,
            sheet.max_column + 1
        ):

            cell = sheet.cell(
                row=row,
                column=col
            )


            # ==================================
            # ЗАЛИВКА
            # ==================================

            if row <= 3:

                # Первые 3 строки голубые
                cell.fill = header_fill

            else:

                # Определяем номер образца
                # A-B = образец 0
                # C-D = образец 1
                # E-F = образец 2
                sample_number = (col - 1) // 2

                # Каждый второй образец серый
                if sample_number % 2 == 1:
                    cell.fill = alternate_fill


            # ==================================
            # ШРИФТ И ВЫРАВНИВАНИЕ
            # ==================================

            if row <= 3:

                cell.font = header_font

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            else:

                # Числа выравниваем вправо
                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )


            # ==================================
            # ВЕРТИКАЛЬНЫЕ ЛИНИИ
            # ==================================

            # Каждый файл занимает два столбца.
            #
            # После B, D, F, H...
            # делаем жирную линию.
            #
            # После A, C, E, G...
            # обычную тонкую.

            if (
                col % 2 == 0
                and col < sheet.max_column
            ):

                right_border = thick_line

            else:

                right_border = thin_line


            # ==================================
            # ГОРИЗОНТАЛЬНЫЕ ЛИНИИ
            # ==================================

            bottom_border = thin_line


            # ==================================
            # ПРИМЕНЯЕМ ГРАНИЦЫ
            # ==================================

            cell.border = Border(
                right=right_border,
                bottom=bottom_border
            )


    # ==========================================
    # ЗАКРЕПЛЯЕМ ПЕРВЫЕ 3 СТРОКИ
    # ==========================================

    sheet.freeze_panes = "A4"


    # ==========================================
    # ФОРМАТ ЧИСЕЛ
    # ==========================================

    for row in sheet.iter_rows(
        min_row=4,
        max_row=sheet.max_row,
        min_col=1,
        max_col=sheet.max_column
    ):

        for cell in row:

            if isinstance(
                cell.value,
                (int, float)
            ):

                cell.number_format = "0.########"


    # ==========================================
    # АВТОМАТИЧЕСКАЯ ШИРИНА СТОЛБЦОВ
    # ==========================================

    for col in range(
        1,
        sheet.max_column + 1
    ):

        max_length = 0

        for row in range(
            1,
            sheet.max_row + 1
        ):

            value = sheet.cell(
                row=row,
                column=col
            ).value

            if value is not None:

                value_length = len(
                    str(value)
                )

                if value_length > max_length:
                    max_length = value_length


        column_letter = get_column_letter(col)

        # Минимум 12, максимум 24
        width = max(
            12,
            min(
                max_length + 2,
                24
            )
        )

        sheet.column_dimensions[
            column_letter
        ].width = width


    # ==========================================
    # ВЫСОТА ВЕРХНИХ СТРОК
    # ==========================================

    sheet.row_dimensions[1].height = 22
    sheet.row_dimensions[2].height = 22
    sheet.row_dimensions[3].height = 24

    # ==========================================
    # СОХРАНЕНИЕ
    # ==========================================

    results_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    output_file = (
        results_folder
        / f"{experiment}_processed.xlsx"
    )


    try:

        workbook.save(output_file)

    except PermissionError:

        raise PermissionError(
            f"Не удалось сохранить файл:\n"
            f"{output_file.name}\n\n"
            f"Скорее всего Excel-файл сейчас открыт.\n"
            f"Закройте его и повторите обработку."
        )


    return output_file

# ==========================================
# ГЛАВНОЕ ОКНО
# ==========================================

class MPTProcessorApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("MPT PROCESSOR")
        self.geometry("1450x900")
        self.minsize(1200, 750)

        self.program_folder = get_program_folder()
        self.input_folder = self.program_folder / "MPT_files"
        self.results_folder = self.program_folder / "Results"

        self.input_folder.mkdir(exist_ok=True)
        self.results_folder.mkdir(exist_ok=True)

        self.mpt_files = []
        self.mass_vars = {}
        self.mass_entries = {}

        self.experiment_var = tk.StringVar(value="LSV")
        self.x_name_var = tk.StringVar()
        self.y_name_var = tk.StringVar()
        self.x_unit_var = tk.StringVar()
        self.y_unit_var = tk.StringVar()
        self.files_count_var = tk.StringVar(value="Найдено .mpt файлов: 0")
        self.lsv_count_var = tk.StringVar(value="LSV: 0")
        self.cstv_count_var = tk.StringVar(value="CstV: 0")
        self.gcpl_count_var = tk.StringVar(value="GCPL: 0")
        self.ocv_count_var = tk.StringVar(value="OCV: 0")
        self.status_var = tk.StringVar(value="Готово к работе")

        self.build_ui()
        self.refresh_files()
        self.update_experiment_settings()
        

    # ==========================================
    # СОЗДАНИЕ ИНТЕРФЕЙСА
    # ==========================================
    def build_ui(self):
        # ===== Верхний заголовок =====
        header_frame = tk.Frame(self, padx=20, pady=15)
        header_frame.pack(fill="x")

        logo_label = tk.Label(
            header_frame,
            text="∿",
            font=("Segoe UI", 34, "bold"),
            fg="#048D92"
        )
        logo_label.pack(side="left", padx=(0, 10))

        titles_frame = tk.Frame(header_frame)
        titles_frame.pack(side="left")

        tk.Label(
            titles_frame,
            text="MPT PROCESSOR",
            font=("Segoe UI", 28, "bold")
        ).pack(anchor="w")

        tk.Label(
            titles_frame,
            text="Created special for ANYNAME.",
            font=("Segoe UI", 14),
            fg="gray40"
        ).pack(anchor="w")

        # ===== Основная область =====
        content_frame = tk.Frame(self, padx=15, pady=5)
        content_frame.pack(fill="both", expand=True)

        left_frame = tk.Frame(content_frame)
        left_frame.pack(side="left", fill="y", padx=(0, 10))

        right_frame = tk.Frame(content_frame)
        right_frame.pack(side="left", fill="both", expand=True)

        # ==========================================
        # 1. ПАПКИ
        # ==========================================
        folders_frame = ttk.LabelFrame(left_frame, text="1. Папки (статические)", padding=15)
        folders_frame.pack(fill="x", pady=(0, 10))

        row1 = tk.Frame(folders_frame)
        row1.pack(fill="x", pady=5)

        tk.Label(row1, text="Папка с исходными файлами:", font=("Segoe UI", 11)).pack(side="left")
        tk.Label(row1, text="MPT_files", font=("Segoe UI", 11, "bold"), fg="#1E5EFF").pack(side="left", padx=6)

        ttk.Button(
            row1,
            text="ОТКРЫТЬ",
            command=lambda: open_folder(self.input_folder)
        ).pack(side="right")

        row2 = tk.Frame(folders_frame)
        row2.pack(fill="x", pady=5)

        tk.Label(row2, text="Папка результатов:", font=("Segoe UI", 11)).pack(side="left")
        tk.Label(row2, text="Results", font=("Segoe UI", 11, "bold"), fg="green4").pack(side="left", padx=6)

        ttk.Button(
            row2,
            text="ОТКРЫТЬ",
            command=lambda: open_folder(self.results_folder)
        ).pack(side="right")

         # ==========================================
        # ОБЩЕЕ КОЛИЧЕСТВО ФАЙЛОВ
        # ==========================================

        row3 = tk.Frame(folders_frame)
        row3.pack(fill="x", pady=(10, 0))

        tk.Label(
            row3,
            textvariable=self.files_count_var,
            font=("Segoe UI", 11, "bold"),
            fg="#1E5EFF"
        ).pack(side="left")

        ttk.Button(
            row3,
            text="ОБНОВИТЬ",
            command=self.refresh_files
        ).pack(side="right")


        # ==========================================
        # КОЛИЧЕСТВО ФАЙЛОВ ПО ЭКСПЕРИМЕНТАМ
        # ==========================================

        experiments_count_frame = tk.Frame(
            folders_frame
        )

        experiments_count_frame.pack(
            fill="x",
            pady=(10, 0)
        )


        tk.Label(
            experiments_count_frame,
            textvariable=self.lsv_count_var,
            font=("Segoe UI", 10)
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 30)
        )


        tk.Label(
            experiments_count_frame,
            textvariable=self.cstv_count_var,
            font=("Segoe UI", 10)
        ).grid(
            row=0,
            column=1,
            sticky="w"
        )


        tk.Label(
            experiments_count_frame,
            textvariable=self.gcpl_count_var,
            font=("Segoe UI", 10)
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 30),
            pady=(5, 0)
        )


        tk.Label(
            experiments_count_frame,
            textvariable=self.ocv_count_var,
            font=("Segoe UI", 10)
        ).grid(
            row=1,
            column=1,
            sticky="w",
            pady=(5, 0)
        )

            # ==========================================
        # 2. ТИП ЭКСПЕРИМЕНТА
        # ==========================================

        experiment_frame = ttk.LabelFrame(
            left_frame,
            text="2. Тип эксперимента",
            padding=15
        )

        experiment_frame.pack(
            fill="x",
            pady=(0, 10)
        )

        self.experiment_combo = ttk.Combobox(
            experiment_frame,
            textvariable=self.experiment_var,
            values=[
                "LSV",
                "CstV",
                "GCPL",
                "OCV"
            ],
            state="readonly",
            font=("Segoe UI", 12)
        )

        self.experiment_combo.pack(
            fill="x"
        )

        self.experiment_combo.bind(
            "<<ComboboxSelected>>",
            self.on_experiment_change
        )

        # ==========================================
        # 3. ПАРАМЕТРЫ ВЫВОДА
        # ==========================================
        output_frame = ttk.LabelFrame(left_frame, text="3. Параметры вывода", padding=15)
        output_frame.pack(fill="x", pady=(0, 10))

        axes_frame = tk.Frame(output_frame)
        axes_frame.pack(fill="x")

        # ---- Ось X ----
        x_frame = ttk.LabelFrame(axes_frame, text="Ось X", padding=12)
        x_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        tk.Label(x_frame, text="Переменная", font=("Segoe UI", 10)).pack(anchor="w")
        self.x_name_entry = ttk.Entry(x_frame, textvariable=self.x_name_var, state="readonly")
        self.x_name_entry.pack(fill="x", pady=(3, 10))

        tk.Label(x_frame, text="Единицы", font=("Segoe UI", 10)).pack(anchor="w")
        self.x_unit_combo = ttk.Combobox(
            x_frame,
            textvariable=self.x_unit_var,
            state="readonly"
        )
        self.x_unit_combo.pack(fill="x", pady=(3, 0))

        # ---- Ось Y ----
        y_frame = ttk.LabelFrame(axes_frame, text="Ось Y", padding=12)
        y_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        tk.Label(y_frame, text="Переменная", font=("Segoe UI", 10)).pack(anchor="w")
        self.y_name_entry = ttk.Entry(y_frame, textvariable=self.y_name_var, state="readonly")
        self.y_name_entry.pack(fill="x", pady=(3, 10))

        tk.Label(y_frame, text="Единицы", font=("Segoe UI", 10)).pack(anchor="w")
        self.y_unit_combo = ttk.Combobox(
            y_frame,
            textvariable=self.y_unit_var,
            state="readonly"
        )
        self.y_unit_combo.pack(fill="x", pady=(3, 0))
        self.y_unit_combo.bind("<<ComboboxSelected>>", self.on_y_unit_change)

        # ==========================================
        # 4. МАССА АКТИВНОГО МАТЕРИАЛА
        # ==========================================
        self.mass_frame = ttk.LabelFrame(left_frame, text="4. Масса активного материала", padding=15)
        self.mass_frame.pack(fill="x", pady=(0, 10))

        header_row = tk.Frame(self.mass_frame)
        header_row.pack(fill="x", pady=(0, 5))

        tk.Label(header_row, text="Ячейка", width=20, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(header_row, text="Масса", width=12, font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w")
        tk.Label(header_row, text="", width=6, font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w")

        self.mass_rows_frame = tk.Frame(self.mass_frame)
        self.mass_rows_frame.pack(fill="x")

        self.mass_note_label = tk.Label(
            self.mass_frame,
            text="Требуется только для нормирования тока (mA/g или A/g)",
            font=("Segoe UI", 9),
            fg="gray40"
        )
        self.mass_note_label.pack(anchor="w", pady=(8, 0))

        # ==========================================
        # КНОПКИ
        # ==========================================
        buttons_frame = tk.Frame(left_frame)
        buttons_frame.pack(fill="x", pady=(0, 10))

        self.process_button = tk.Button(
            buttons_frame,
            text="▶ ОБРАБОТАТЬ",
            font=("Segoe UI", 12, "bold"),
            bg="#1E5EFF",
            fg="white",
            padx=10,
            pady=12,
            command=self.process_files
        )
        self.process_button.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.results_button = ttk.Button(
            buttons_frame,
            text="ОТКРЫТЬ RESULTS",
            command=lambda: open_folder(self.results_folder)
        )
        self.results_button.pack(side="left", fill="x", expand=True, padx=(6, 0), ipady=12)

        # ==========================================
        # ПРАВАЯ ЧАСТЬ - ПРЕДПРОСМОТР
        # ==========================================
        preview_header = tk.Frame(right_frame)
        preview_header.pack(fill="x", pady=(0, 8))

        tk.Label(
            preview_header,
            text="Предпросмотр Excel",
            font=("Segoe UI", 16, "bold"),
            fg="#1E5EFF"
        ).pack(side="left")


        preview_frame = ttk.LabelFrame(right_frame, padding=10)
        preview_frame.pack(fill="both", expand=True)

        columns = ("row", "c1", "c2", "c3", "c4", "c5", "c6")
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=25)

        self.preview_tree.heading("row", text="#")
        self.preview_tree.heading("c1", text="")
        self.preview_tree.heading("c2", text="")
        self.preview_tree.heading("c3", text="")
        self.preview_tree.heading("c4", text="")
        self.preview_tree.heading("c5", text="")
        self.preview_tree.heading("c6", text="")

        self.preview_tree.column("row", width=45, anchor="center")
        for col in ("c1", "c2", "c3", "c4", "c5", "c6"):
            self.preview_tree.column(col, width=130, anchor="center")

        scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)

        self.preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.preview_tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")

        # ==========================================
        # СТАТУС БАР
        # ==========================================
        status_frame = tk.Frame(self, bd=1, relief="sunken", padx=10, pady=8)
        status_frame.pack(fill="x", side="bottom")

        tk.Label(
            status_frame,
            text="●",
            fg="#1E5EFF",
            font=("Segoe UI", 14, "bold")
        ).pack(side="left", padx=(0, 6))

        tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 11)
        ).pack(side="left")

    def refresh_files(self):

    # ==========================================
    # ИЩЕМ ВСЕ .MPT
    # ==========================================

        self.mpt_files = sorted(
            self.input_folder.glob("*.mpt")
    )


    # ==========================================
    # ОБЩЕЕ КОЛИЧЕСТВО
    # ==========================================

        self.files_count_var.set(
            f"Найдено .mpt файлов: {len(self.mpt_files)}"
    )


    # ==========================================
    # СЧИТАЕМ ТИПЫ ЭКСПЕРИМЕНТОВ
    # ==========================================

        lsv_count = 0
        cstv_count = 0
        gcpl_count = 0
        ocv_count = 0


        for file in self.mpt_files:

            file_name = file.name.upper()

            if "_LSV_" in file_name:
                lsv_count += 1

            elif "_CSTV_" in file_name:
                cstv_count += 1

            elif "_GCPL_" in file_name:
                gcpl_count += 1

            elif "_OCV_" in file_name:
                ocv_count += 1


    # ==========================================
    # ПОКАЗЫВАЕМ РЕЗУЛЬТАТ
    # ==========================================

        self.lsv_count_var.set(
            f"LSV: {lsv_count}"
    )

        self.cstv_count_var.set(
            f"CstV: {cstv_count}"
    )

        self.gcpl_count_var.set(
            f"GCPL: {gcpl_count}"
    )

        self.ocv_count_var.set(
            f"OCV: {ocv_count}"
    )


    # ==========================================
    # ОБНОВЛЯЕМ ОСТАЛЬНОЙ ИНТЕРФЕЙС
    # ==========================================

        self.rebuild_mass_rows()
        self.update_status_ready()

    # ==========================================
    # ИЗМЕНЕНИЕ ТИПА ЭКСПЕРИМЕНТА
    # ==========================================
    def on_experiment_change(self, event=None):
        self.update_experiment_settings()
        self.rebuild_mass_rows()
        self.update_status_ready()

    # ==========================================
    # НАСТРОЙКИ ПО ЭКСПЕРИМЕНТУ
    # ==========================================
    def update_experiment_settings(self):
        experiment = self.experiment_var.get()

        if experiment == "LSV":
            x_name = "Ewe"
            x_units = ["V", "mV"]
            y_name = "<I>"
            y_units = ["mA", "A", "mA/g", "A/g"]

        elif experiment == "CstV":
            x_name = "time"
            x_units = ["s", "min", "h"]
            y_name = "<I>"
            y_units = ["mA", "A", "mA/g", "A/g"]

        elif experiment == "GCPL":
            x_name = "time"
            x_units = ["s", "min", "h"]
            y_name = "Ewe"
            y_units = ["V", "mV"]

        else:  # OCV
            x_name = "time"
            x_units = ["s", "min", "h"]
            y_name = "Ewe"
            y_units = ["V", "mV"]

        self.x_name_var.set(x_name)
        self.y_name_var.set(y_name)

        self.x_unit_combo["values"] = x_units
        self.y_unit_combo["values"] = y_units

        self.x_unit_var.set(x_units[0])
        self.y_unit_var.set(y_units[0])

        self.update_mass_entries_state()

    # ==========================================
    # НУЖНА ЛИ МАССА
    # ==========================================
    def on_y_unit_change(self, event=None):
        self.update_mass_entries_state()

    def update_mass_entries_state(self):
        need_mass = self.need_mass()

        for entry in self.mass_entries.values():
            if need_mass:
                entry.configure(state="normal")
            else:
                entry.configure(state="disabled")

    def need_mass(self):
        experiment = self.experiment_var.get()
        y_unit = self.y_unit_var.get()

        if experiment in ["LSV", "CstV"] and y_unit in ["mA/g", "A/g"]:
            return True
        return False

    # ==========================================
    # ПЕРЕСТРОИТЬ ТАБЛИЦУ МАСС
    # ==========================================
    def rebuild_mass_rows(self):
        for widget in self.mass_rows_frame.winfo_children():
            widget.destroy()

        self.mass_entries.clear()

        selected_experiment = self.experiment_var.get()
        marker = f"_{selected_experiment}_"

        selected_files = [file for file in self.mpt_files if marker in file.name]

        cell_names = []
        for file in selected_files:
            cell_name = extract_cell_name(file)
            if cell_name not in cell_names:
                cell_names.append(cell_name)

        if not cell_names:
            label = tk.Label(
                self.mass_rows_frame,
                text="Нет файлов выбранного типа в папке MPT_files",
                font=("Segoe UI", 10),
                fg="gray40"
            )
            label.pack(anchor="w", pady=5)
            return

        for index, cell_name in enumerate(cell_names):
            row = tk.Frame(self.mass_rows_frame)
            row.pack(fill="x", pady=2)

            tk.Label(row, text=cell_name, width=20, anchor="w", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")

            if cell_name not in self.mass_vars:
                self.mass_vars[cell_name] = tk.StringVar()

            entry = ttk.Entry(row, textvariable=self.mass_vars[cell_name], width=12)
            entry.grid(row=0, column=1, padx=5, sticky="w")

            tk.Label(row, text="mg", width=6, anchor="w", font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w")

            self.mass_entries[cell_name] = entry

        self.update_mass_entries_state()

    # ==========================================
    # ДЕМО-ПРЕДПРОСМОТР
    # ==========================================
    def show_real_preview(
    self,
    processed_data,
    experiment,
    x_unit,
    y_unit
):

    # Очищаем старый предпросмотр
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

    # Пока предпросмотр показывает максимум 3 образца.
    # Сам Excel позже сможет содержать хоть 50.
        preview_data = processed_data[:3]

    # Названия величин
        if experiment == "LSV":
            x_name = "Ewe"
            y_name = "<I>"

        elif experiment == "CstV":
            x_name = "time"
            y_name = "<I>"

        elif experiment in ["GCPL", "OCV"]:
            x_name = "time"
            y_name = "Ewe"

        else:
            x_name = "X"
            y_name = "Y"

    # --------------------------
    # Строка 1 — величины
    # --------------------------

        row1 = ["1"]

        for sample in preview_data:
            row1.append(x_name)
            row1.append(y_name)

        while len(row1) < 7:
            row1.append("")

        self.preview_tree.insert(
        "",
        "end",
        values=row1
    )

    # --------------------------
    # Строка 2 — единицы
    # --------------------------

        row2 = ["2"]

        for sample in preview_data:
            row2.append(x_unit)
            row2.append(y_unit)

        while len(row2) < 7:
            row2.append("")

        self.preview_tree.insert(
        "",
        "end",
        values=row2
    )

    # --------------------------
    # Строка 3 — Cell name
    # --------------------------

        row3 = ["3"]

        for sample in preview_data:

            cell_name = sample["cell_name"]

            row3.append(cell_name)
            row3.append(cell_name)

        while len(row3) < 7:
            row3.append("")

        self.preview_tree.insert(
        "",
        "end",
        values=row3
    )

    # --------------------------
    # Строки данных
    # --------------------------

        max_rows = 20

        longest_sample = max(
            len(sample["x"])
            for sample in preview_data
    )

        rows_to_show = min(
           longest_sample,
           max_rows
    )

        for index in range(rows_to_show):

            row = [str(index + 4)]

            for sample in preview_data:

                x = sample["x"]
                y = sample["y"]

                if index < len(x):
                    x_value = x.iloc[index]
                    row.append(f"{x_value:.8g}")
                else:
                    row.append("")

                if index < len(y):
                    y_value = y.iloc[index]
                    row.append(f"{y_value:.8g}")
                else:
                    row.append("")

            while len(row) < 7:
                row.append("")

            self.preview_tree.insert(
            "",
            "end",
            values=row
        )

    # ==========================================
    # СТАТУС
    # ==========================================
    def update_status_ready(self):
        experiment = self.experiment_var.get()
        self.status_var.set(f"Готово к обработке выбранных {experiment} файлов")

    # ==========================================
    # ПОКА ЗАГЛУШКИ
    # ==========================================

    def process_files(self):

        experiment = self.experiment_var.get()
        x_unit = self.x_unit_var.get()
        y_unit = self.y_unit_var.get()

        marker = f"_{experiment}_"

        selected_files = [
            file for file in self.mpt_files
            if marker in file.name
    ]

        if len(selected_files) == 0:
            messagebox.showwarning("Нет файлов",f"В папке MPT_files нет файлов типа {experiment}.")
            return

        processed_data = []

        processed_data = []

        try:

            for file in selected_files:

                cell_name = extract_cell_name(file)

                # ==========================
                # МАССА
                # ==========================

                mass_mg = None

                if self.need_mass():

                    mass_text = self.mass_vars[cell_name].get().strip()

                    if mass_text == "":
                        raise ValueError(
                            f"Не указана масса для {cell_name}"
                        )

                    mass_text = mass_text.replace(",", ".")
                    mass_mg = float(mass_text)

                    if mass_mg <= 0:
                        raise ValueError(
                            f"Масса {cell_name} должна быть больше нуля."
                        )


                # ==========================
                # ЧИТАЕМ .MPT
                # ==========================

                df = read_mpt(file)


                # ==========================
                # ПОЛУЧАЕМ X И Y
                # ==========================

                x, y = prepare_experiment_data(
                    df=df,
                    experiment=experiment,
                    x_unit=x_unit,
                    y_unit=y_unit,
                    mass_mg=mass_mg
                )


                # ==========================
                # ДОБАВЛЯЕМ ЭТОТ ОБРАЗЕЦ
                # В ОБЩИЙ СПИСОК
                # ==========================

                processed_data.append(
                    {
                        "cell_name": cell_name,
                        "x": x.reset_index(drop=True),
                        "y": y.reset_index(drop=True)
                    }
                )


        except Exception as error:

            messagebox.showerror(
                "Ошибка обработки",
                str(error)
            )

            return

        except Exception as error:

            messagebox.showerror(
            "Ошибка обработки",
            str(error)
        )

            return

    # Показываем реальные данные
        self.show_real_preview(
        processed_data,
        experiment,
        x_unit,
        y_unit
    )

  # --------------------------
        # Названия величин
        # --------------------------

        if experiment == "LSV":

            x_name = "Ewe"
            y_name = "<I>"

        elif experiment == "CstV":

            x_name = "time"
            y_name = "<I>"

        elif experiment in ["GCPL", "OCV"]:

            x_name = "time"
            y_name = "Ewe"


        # --------------------------
        # Создаём Excel
        # --------------------------

        try:

            output_file = create_excel(
                processed_data=processed_data,
                experiment=experiment,
                x_name=x_name,
                y_name=y_name,
                x_unit=x_unit,
                y_unit=y_unit,
                results_folder=self.results_folder
            )

        except PermissionError as error:

            messagebox.showerror(
                "Не удалось сохранить Excel",
                str(error)
            )

            return

        except Exception as error:

            messagebox.showerror(
                "Ошибка создания Excel",
                str(error)
            )

            return


        # Если всё прошло успешно
        messagebox.showinfo(
            "Готово!",
            f"Excel успешно создан:\n\n"
            f"{output_file.name}\n\n"
            f"Файл находится в папке Results."
        )

# ==========================================
# ЗАПУСК
# ==========================================

if __name__ == "__main__":
    app = MPTProcessorApp()
    app.mainloop()