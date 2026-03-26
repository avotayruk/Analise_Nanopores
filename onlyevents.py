import sys
import os
from openpyxl import Workbook

# Важно для PyInstaller
if getattr(sys, 'frozen', False):
    # Если запущено как собранное приложение
    import matplotlib

    matplotlib.use('TkAgg')  # Явно указываем бэкенд

    # Добавляем путь к ресурсам PyInstaller
    base_path = sys._MEIPASS
else:
    # Если запущено как скрипт Python
    base_path = os.path.dirname(os.path.abspath(__file__))

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
import tkinter as tk
from tkinter import messagebox
import pandas as pd

from defs import calculate_triggers
from defs import EMA_calculate_baseline
from defs import calculate_detecting_down
from defs import calculate_detecting_all
from defs import creating_table
from defs import filtering
from defs import calculation_one
from defs import calculation_both
from defs import save_summary_excel
from defs import save_raw_events_excel

from matplotlib.ticker import MultipleLocator
import openpyxl

# =========================
# 1. Чтение данных
# =========================

def get_parameters_gui():


    result = {}

    def on_run():
        try:
            result['fs_khz'] = float(entry_fs.get())
            result['a'] = float(entry_a.get())
            result['k'] = float(entry_k.get())
            result['subsample'] = int(entry_subsample.get())
            result['event_buffer'] = int(entry_event_buffer.get())
            result['t_start'] = float(entry_t_start.get())
            result['t_end'] = float(entry_t_end.get())
            result['filename'] = filename.get().strip()
            result['window_length'] = int(window_length.get())
            result['polyorder'] = int(polyorder.get())
            result['METOD'] = METOD.get().strip()
            result['positive_events'] = positive_events_var.get()
        except ValueError:
            messagebox.showerror("Ошибка", "Ну чтото не так")
            return

        root.quit()   # ВАЖНО: выйти из mainloop

    root = tk.Tk()
    root.title("Параметры анализа нанопорного сигнала")
    root.geometry("420x400")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=15, pady=15)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Характеристики эксперимента").grid(row=0, column=0, sticky="w")

    tk.Label(frame, text="Имя файла (пример: GD.txt)").grid(row=1, column=0, sticky="w")
    filename = tk.Entry(frame, width=15)
    filename.insert(0, "0DNA500bp.txt")
    filename.grid(row=1, column=1)

    tk.Label(frame, text="Частота дискретизации (кГц):").grid(row=2, column=0, sticky="w")
    entry_fs = tk.Entry(frame, width=15)
    entry_fs.insert(0, "50")
    entry_fs.grid(row=2, column=1)

    tk.Label(frame, text="Начало анализа, с:").grid(row=3, column=0, sticky="w")
    entry_t_start = tk.Entry(frame, width=15)
    entry_t_start.insert(0, "0")
    entry_t_start.grid(row=3, column=1)

    tk.Label(frame, text="Конец анализа, с:").grid(row=4, column=0, sticky="w")
    entry_t_end = tk.Entry(frame, width=15)
    entry_t_end.insert(0, "-1")
    entry_t_end.grid(row=4, column=1)

    tk.Label(frame, text="Коэффициент триггера k:").grid(row=5, column=0, sticky="w")
    entry_k = tk.Entry(frame, width=15)
    entry_k.insert(0, "4.5")
    entry_k.grid(row=5, column=1)


    tk.Label(frame, text="Points of Window").grid(row=7, column=0, sticky="w")
    window_length = tk.Entry(frame, width=15)
    window_length.insert(0, "8000")
    window_length.grid(row=7, column=1)

    tk.Label(frame, text="polyorder").grid(row=8, column=0, sticky="w")
    polyorder = tk.Entry(frame, width=15)
    polyorder.insert(0, "2")
    polyorder.grid(row=8, column=1)

    tk.Label(frame, text="Метод EMA baseline").grid(row=9, column=0, sticky="w")

    tk.Label(frame, text="EMA baseline a:").grid(row=10, column=0, sticky="w")
    entry_a = tk.Entry(frame, width=15)
    entry_a.insert(0, "0.9995")
    entry_a.grid(row=10, column=1)

    tk.Label(frame, text="Вид графика").grid(row=11, column=0, sticky="w")

    tk.Label(frame, text="Прореживание (subsample):").grid(row=12, column=0, sticky="w")
    entry_subsample = tk.Entry(frame, width=15)
    entry_subsample.insert(0, "100")
    entry_subsample.grid(row=12, column=1)

    tk.Label(frame, text="Кол-во точек в окресности события:").grid(row=13, column=0, sticky="w")
    entry_event_buffer = tk.Entry(frame, width=15)
    entry_event_buffer.insert(0, "200")
    entry_event_buffer.grid(row=13, column=1)

    tk.Label(frame, text="Метод  событий (SG или EMA)").grid(row=14, column=0, sticky="w")
    METOD = tk.Entry(frame, width=15)
    METOD.insert(0, "SG и EMA")
    METOD.grid(row=14, column=1)

    tk.Label(frame, text="Считать положительные события?").grid(row=15, column=0, sticky="w")
    positive_events_var = tk.IntVar()
    positive_events_cb = tk.Checkbutton(frame, variable=positive_events_var, onvalue=1, offvalue=0)
    positive_events_cb.grid(row=15, column=1, sticky="w")



    tk.Button(frame, text="Run analysis", command=on_run, width=20)\
        .grid(row=17, column=0, columnspan=2, pady=20)
    root.bind('<Return>', lambda event: on_run())


    root.mainloop()
    root.destroy()

    if not result:
        raise RuntimeError("Ввод параметров отменён")

    return (
        result['filename'],
        result['fs_khz'],
        result['t_start'],
        result['t_end'],
        result['k'],
        result['window_length'],
        result['polyorder'],
        result['a'],
        result['subsample'],
        result['event_buffer'],
        result['METOD'],
        result['positive_events']
    )
filename, fs_khz, t_start, t_end, k, window_length, polyorder, a, subsample, event_buffer, METOD, positive_events = get_parameters_gui()

fs = fs_khz * 1e3           # Гц
dt = 1.0 / fs               # шаг времени (с)
window = 200
symmetry_ratio = 0.5


with open(filename, 'r') as f:
    data = [line.strip().replace(',', '.') for line in f]

values = np.array(data, dtype=float)


n_points = len(values)
time = np.arange(n_points) * dt


if t_end < 0:
    t_end = time[-1]

mask = (time >= t_start) & (time <= t_end)

time = time[mask]
values = values[mask]

n_points = len(values)

print(f"Прочитано точек: {n_points}")
print(f'Среднее значение изначального сигнала: {np.mean(values)}')
# =========================
# 2. Вычисление baseline (EMA)
# =========================



if METOD in ["SG", "EMA"]:
    (std_value, trigger_line, trigger, raw_events,
     filtered_events, negative_count, positive_count, delta_I) = calculation_one(values, a,
                                                                        k, positive_events, n_points, window,
                                                                        symmetry_ratio, dt, METOD,
                                                                        window_length, polyorder)
else:
    time_ema = time
    values_ema = values
    (std_value, trigger_line, trigger, raw_events,
     filtered_events, negative_count, positive_count,
     ema_std_value, ema_trigger_line, ema_trigger, ema_raw_events,
     ema_filtered_events, ema_negative_count, ema_positive_count, delta_I, ema_delta_I) = calculation_both(values, k, positive_events,
                                                                                     n_points, window, symmetry_ratio,
                                                                                     dt, METOD, values_ema, a, window_length, polyorder)


#===============================================================================
#===============================================================================

# =========================
# 6. Формирование таблицы событий
# =========================

if METOD in ["SG", "EMA"]:

    events_table = []
    events_table_sorted = creating_table(filtered_events, delta_I, time, events_table, dt)

else:

    events_table = []
    events_table_sorted = creating_table(filtered_events, delta_I, time, events_table, dt)
    ema_events_table = []
    ema_events_table_sorted = creating_table(ema_filtered_events, ema_delta_I, time, ema_events_table, dt)

#============================================================
#============================================================

# =========================
# 7. Визуализация сигнала
# =========================



# Получаем длительности событий в миллисекундах
# Точечная диаграмма по амплитуде
plt.figure(num='Распределение по амплитуде', figsize=(10, 5))

if METOD in ["SG", "EMA"]:

    sg_neg_durations = []
    sg_neg_amplitudes = []
    sg_pos_durations = []
    sg_pos_amplitudes = []

    for idx, (start, end) in enumerate(filtered_events):
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)
        duration = (end - start + 1) * dt * 1000  # в мс
        amplitude = np.min(delta_I[start:end + 1]) if event_mean < 0 else np.max(delta_I[start:end + 1])

        if event_mean < 0:  # Отрицательное SG событие

            sg_neg_durations.append(duration)
            sg_neg_amplitudes.append(amplitude)

        else:  # Положительное SG событие

            sg_pos_durations.append(duration)
            sg_pos_amplitudes.append(amplitude)

else:
    sg_neg_durations = []
    sg_neg_amplitudes = []
    sg_pos_durations = []
    sg_pos_amplitudes = []

    for idx, (start, end) in enumerate(filtered_events):
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)
        duration = (end - start + 1) * dt * 1000  # в мс
        amplitude = np.min(delta_I[start:end + 1]) if event_mean < 0 else np.max(delta_I[start:end + 1])

        if event_mean < 0:  # Отрицательное SG событие

            sg_neg_durations.append(duration)
            sg_neg_amplitudes.append(amplitude)

        else:  # Положительное SG событие

            sg_pos_durations.append(duration)
            sg_pos_amplitudes.append(amplitude)

    ema_neg_durations = []
    ema_neg_amplitudes = []
    ema_pos_durations = []
    ema_pos_amplitudes = []

    for idx, (start, end) in enumerate(ema_filtered_events):
        event_segment = ema_delta_I[start:end + 1]
        event_mean = np.mean(event_segment)
        duration = (end - start + 1) * dt * 1000  # в мс
        amplitude = np.min(ema_delta_I[start:end + 1]) if event_mean < 0 else np.max(ema_delta_I[start:end + 1])

        if event_mean < 0:  # Отрицательное EMA событие

            ema_neg_durations.append(duration)
            ema_neg_amplitudes.append(amplitude)

        else:  # Положительное EMA событие

            ema_pos_durations.append(duration)
            ema_pos_amplitudes.append(amplitude)


if METOD in ["SG", "EMA"]:
    time_sub = time[::subsample]
    delta_sub = delta_I[::subsample]
else:
    time_sub = time[::subsample]
    delta_sub = delta_I[::subsample]
    ema_delta_sub = ema_delta_I[::subsample]
    values_sub = values[::subsample]


plt.figure(num='События на графике', figsize=(20, 10))

# 7.1 Основной сигнал

if METOD in ["SG", "EMA"]:
    plt.plot(time_sub, delta_sub, color='lightblue')
else:
    plt.plot(time_sub, ema_delta_sub, color='yellow')
    plt.plot(time_sub, delta_sub, color='lightblue')

# 7.2 Зеленые области событий + полный сигнал вокруг события
if METOD in ["SG", "EMA"]:
    for start, end in raw_events:
        seg_start = max(0, start - event_buffer)
        seg_end = min(n_points - 1, end + event_buffer)

        # Серая область события
        plt.axvspan(time[start], time[end], color='gray', alpha=0.15)

        # Сигнал вокруг события (тонкий и прозрачный)
        plt.plot(time[seg_start:seg_end + 1],
                 delta_I[seg_start:seg_end + 1],
                 color='gray',
                 alpha=0.4,
                 linewidth=0.5)
    for start, end in filtered_events:
        seg_start = max(0, start - event_buffer)
        seg_end = min(n_points - 1, end + event_buffer)

        # Определяем знак события
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)

        if event_mean < 0:  # Отрицательное событие - синий
            color = 'blue'
            alpha = 0.2
        else:  # Положительное событие - красный
            color = 'red'
            alpha = 0.2

        # Область события
        plt.axvspan(time[start], time[end], color=color, alpha=alpha)

        # Полный сигнал вокруг события
        plt.plot(time[seg_start:seg_end + 1], delta_I[seg_start:seg_end + 1],
                 color=color, alpha=0.8, linewidth=0.8)

        legend_elements = [
            Patch(facecolor='blue', alpha=0.3, label=f'{METOD} отрицательные'),
            Patch(facecolor='red', alpha=0.3, label=f'{METOD} положительные'),
            Line2D([0], [0], color='blue', linestyle='--', label=f'Trigger {METOD} = {trigger_line:.6f}'),
            Line2D([0], [0], color='lightblue', linestyle='-', label=f'{METOD} (каждая {subsample}-я точка)')

        ]

    plt.axhline(trigger_line, color='blue', linestyle='--')
    plt.axhline(trigger, color='blue', linestyle='--')
#
# if METOD == "EMA":
#
#     for start, end in ema_filtered_events:
#         seg_start = max(0, start - event_buffer)
#         seg_end = min(n_points - 1, end + event_buffer)
#
#         # Определяем знак события
#         event_segment = ema_delta_I[start:end + 1]
#         event_mean = np.mean(event_segment)
#
#         if event_mean < 0:  # Отрицательное событие - фиолетовый
#             color = 'purple'
#             alpha = 0.2
#         else:  # Положительное событие - оранжевый
#             color = 'orange'
#             alpha = 0.2
#
#         # Область события
#         plt.axvspan(time[start], time[end], color=color, alpha=alpha)
#
#         # Полный сигнал вокруг события
#         plt.plot(time[seg_start:seg_end + 1], ema_delta_I[seg_start:seg_end + 1],
#                  color=color, alpha=0.8, linewidth=0.8)
#         legend_elements = [
#             Patch(facecolor='purple', alpha=0.3, label='EMA отрицательные'),
#             Patch(facecolor='orange', alpha=0.3, label='EMA положительные'),
#             Line2D([0], [0], color='red', linestyle='--', label=f'Trigger EMA = {ema_trigger_line:.6f}'),
#             Line2D([0], [0], color='yellow', linestyle='-', label=f'EMA (каждая {subsample}-я точка)')
#
#         ]
#
#     plt.axhline(ema_trigger_line, color='red', linestyle='--')
#     plt.axhline(ema_trigger, color='red', linestyle='--')

if METOD in ["SG", "EMA"]:
    # Сначала EMA события
    for start, end in filtered_events:
        seg_start = max(0, start - event_buffer)
        seg_end = min(n_points - 1, end + event_buffer)

        # Определяем знак события для EMA
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)

        if event_mean < 0:  # Отрицательное EMA - фиолетовый
            color = 'purple'
            alpha = 0.2
        else:  # Положительное EMA - оранжевый
            color = 'orange'
            alpha = 0.2

        # Область события EMA
        plt.axvspan(time[start], time[end], color=color, alpha=alpha)

        # Полный сигнал вокруг события EMA
        plt.plot(time[seg_start:seg_end + 1], delta_I[seg_start:seg_end + 1],
                 color=color, alpha=0.8, linewidth=0.8)

        legend_elements = [
            Patch(facecolor='blue', alpha=0.3, label=f'{METOD} отрицательные'),
            Patch(facecolor='orange', alpha=0.3, label=f'{METOD} положительные'),
            Line2D([0], [0], color='blue', linestyle='--', label=f'Trigger {METOD} = {trigger_line:.6f}'),
            Line2D([0], [0], color='lightblue', linestyle='-', label=f'{METOD} (каждая {subsample}-я точка)'),

        ]

    plt.axhline(trigger_line, color='blue', linestyle='--')
    plt.axhline(trigger, color='blue', linestyle='--')


    # Затем SG события
    for start, end in filtered_events:
        seg_start = max(0, start - event_buffer)
        seg_end = min(n_points - 1, end + event_buffer)

        # Определяем знак события для SG
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)

        if event_mean < 0:  # Отрицательное SG - синий
            color = 'blue'
            alpha = 0.2
        else:  # Положительное SG - красный
            color = 'red'
            alpha = 0.2

        # Область события SG
        plt.axvspan(time[start], time[end], color=color, alpha=alpha)

        # Полный сигнал вокруг события SG
        plt.plot(time[seg_start:seg_end + 1], delta_I[seg_start:seg_end + 1],
                 color=color, alpha=0.8, linewidth=0.8)




plt.legend(handles=legend_elements, loc='upper right')
plt.xlabel('Time (s)')
plt.ylabel('pA')
plt.title('Сигнал с детекцией событий')
plt.grid(alpha=0.3)
plt.tight_layout()


plt.minorticks_on()
plt.grid(True, which='major', alpha=1)
plt.grid(True, which='minor', alpha=0.2)

# =========================
# ОКНО НАВИГАЦИИ (работает параллельно с графиком)
# =========================

# =========================
# ОКНО НАВИГАЦИИ (стабильная версия)
# =========================

fig = plt.gcf()
ax = plt.gca()

nav_root = tk.Tk()
nav_root.title("Навигация по графику")
nav_root.geometry("260x200")
nav_root.resizable(False, False)

tk.Label(nav_root, text="Время от (с)").grid(row=0, column=0)
entry_xmin = tk.Entry(nav_root, width=12)
entry_xmin.grid(row=0, column=1)

tk.Label(nav_root, text="Врема до (с)").grid(row=1, column=0)
entry_xmax = tk.Entry(nav_root, width=12)
entry_xmax.grid(row=1, column=1)

tk.Label(nav_root, text="Ток min").grid(row=2, column=0)
entry_ymin = tk.Entry(nav_root, width=12)
entry_ymin.grid(row=2, column=1)

tk.Label(nav_root, text="Ток max").grid(row=3, column=0)
entry_ymax = tk.Entry(nav_root, width=12)
entry_ymax.grid(row=3, column=1)

def apply_limits():
    try:
        xmin = float(entry_xmin.get())
        xmax = float(entry_xmax.get())
        ymin = float(entry_ymin.get())
        ymax = float(entry_ymax.get())

        ax.set_xlim(xmin, xmax)
        ax.set_ylim(ymin, ymax)
        fig.canvas.draw_idle()

    except ValueError:
        messagebox.showerror("Ошибка", "Введите числовые значения")

tk.Button(nav_root, text="Применить", command=apply_limits)\
    .grid(row=4, column=0, columnspan=2, pady=10)
nav_root.bind('<Return>', lambda event: apply_limits())

plt.show(block=False)
nav_root.mainloop()


params_df = pd.DataFrame({
    'Параметр': [
        'Данные',
        'Частота',
        'Время начала',
        'Время конца',
        'Коэффициент триггера',
        'Длина окна',
        'Степень полинома',
        'Коэффициент a',
        "symetry ratio",
        "окно для фильтрации"
    ],
    'Значение': [
        filename,
        fs_khz,
        t_start,
        t_end,
        k,
        window_length,
        polyorder,
        a,
        symmetry_ratio,
        window
    ]
})

SG_table = pd.DataFrame(events_table_sorted)
SG_table.columns = ['№', 'Время начала (с)', 'Длительность (мс)', 'Амплитуда, pA']

if METOD != "SG":
    ema_table = pd.DataFrame(events_table_sorted)
    ema_table.columns = ['№', 'Время начала (с)', 'Длительность (мс)', 'Амплитуда, pA']

if METOD == "SG":

    # 1 файл (сводка)
    save_summary_excel(
        f"{filename}_SG_{fs_khz}kHz_summary.xlsx",
        params_df,
        {"События SG": SG_table}
    )

    # 2 файл (сырые события)
    save_raw_events_excel(
        f"{filename}_SG_{fs_khz}kHz_raw.xlsx",
        params_df,
        {"События SG": (filtered_events, delta_I)},
        time, event_buffer, n_points
    )

elif METOD == "EMA":

    save_summary_excel(
        f"{filename}_EMA_{fs_khz}kHz_summary.xlsx",
        params_df,
        {"События EMA": ema_table}
    )

    save_raw_events_excel(
        f"{filename}_EMA_{fs_khz}kHz_raw.xlsx",
        params_df,
        {"События EMA": (filtered_events, delta_I)},
        time, event_buffer, n_points
    )


elif METOD == "SG и EMA":

    # 1 файл — ОБЕ ТАБЛИЦЫ
    save_summary_excel(
        f"{filename}_{fs_khz}kHz_summary.xlsx",
        params_df,
        {
            "События SG": SG_table,
            "События EMA": ema_table
        }
    )

    # 2 файл — ВСЕ СИГНАЛЫ
    save_raw_events_excel(
        f"{filename}_{fs_khz}kHz_raw_±{event_buffer}.xlsx",
        params_df,
        {
            "События SG": (filtered_events, delta_I),
            "События EMA": (ema_filtered_events, ema_delta_I)
        },
        time, event_buffer, n_points
    )


#pyinstaller --onefile onlyevents.py

