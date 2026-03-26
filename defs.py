import numpy as np
from openpyxl import Workbook
from scipy.signal import savgol_filter
import pandas as pd


# считает базовую
def EMA_calculate_baseline(n_points, values_ema, a):

    ema = np.zeros_like(values_ema)
    ema[0] = np.mean(values_ema)

    for i in range(1, n_points):
        ema[i] = a * ema[i - 1] + (1 - a) * values_ema[i]

    ema_delta_I = values_ema - ema

    return ema_delta_I

# Рассчет линии триггеров
def calculate_triggers(delta_I, k):

    std_value = np.std(delta_I)
    trigger = std_value * k
    trigger_line = -trigger

    return std_value, trigger_line, trigger

def calculate_detecting_down(delta_I, trigger_line, n_points):
    below_trigger = delta_I < trigger_line
    raw_events = []

    in_event = False
    start_idx = 0

    for i, flag in enumerate(below_trigger):
        if flag and not in_event:
            in_event = True
            start_idx = i
        elif not flag and in_event:
            in_event = False
            end_idx = i - 1
            raw_events.append((start_idx, end_idx))

    if in_event:
        raw_events.append((start_idx, n_points - 1))

    return raw_events

def calculate_detecting_all(delta_I, trigger_line, trigger, n_points):
    # Условие: сигнал вне диапазона [trigger_line, trigger_high]
    # trigger_line - отрицательное значение (нижняя граница)
    # trigger - положительное значение (верхняя граница)
    outside_range = (delta_I < trigger_line) | (delta_I > trigger)
    raw_events = []

    in_event = False
    start_idx = 0

    for i, flag in enumerate(outside_range):
        if flag and not in_event:
            in_event = True
            start_idx = i
        elif not flag and in_event:
            in_event = False
            end_idx = i - 1
            raw_events.append((start_idx, end_idx))

    if in_event:
        raw_events.append((start_idx, n_points - 1))

    return raw_events


def filtering(raw_events, window, symmetry_ratio, n_points, delta_I, trigger_line, trigger, dt):

    # --- параметры расстояний ---
    min_distance_ms2 = 50
    min_distance_ms1 = 5
    min_distance_points2 = int((min_distance_ms2 / 1000) / dt)
    min_distance_points1 = int((min_distance_ms1 / 1000) / dt)

    raw_events_sorted = sorted(raw_events)

    # --- 1. КЛАСТЕРИЗАЦИЯ (фильтр по расстоянию) ---
    clusters = []
    current_cluster = [raw_events_sorted[0]]

    for i in range(1, len(raw_events_sorted)):
        start2, end2 = raw_events_sorted[i]
        start1, end1 = current_cluster[-1]

        distance = start2 - end1

        if min_distance_points1 < distance < min_distance_points2:
            current_cluster.append((start2, end2))
        else:
            clusters.append(current_cluster)
            current_cluster = [(start2, end2)]

    clusters.append(current_cluster)

    clean_raw_events = [cluster[0] for cluster in clusters if len(cluster) == 1]

    # --- 2. РАСШИРЕНИЕ ДО BASELINE ---
    expanded_events = []

    for start, end in clean_raw_events:

        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)

        # --- отрицательное ---
        if event_mean < 0:
            i = start
            while i > 0 and delta_I[i] < 0:
                i -= 1
            new_start = i + 1

            i = end
            while i < n_points - 1 and delta_I[i] < 0:
                i += 1
            new_end = i - 1

        # --- положительное ---
        else:
            i = start
            while i > 0 and delta_I[i] > 0:
                i -= 1
            new_start = i + 1

            i = end
            while i < n_points - 1 and delta_I[i] > 0:
                i += 1
            new_end = i - 1

        expanded_events.append((new_start, new_end))

    # --- 3. ПРОВЕРКА СИММЕТРИИ (уже после расширения!) ---
    filtered_events = []

    for start, end in expanded_events:

        # окно теперь вокруг ВСЕГО события
        w_start = max(0, start - window)
        w_end = min(n_points - 1, end + window)

        segment = delta_I[w_start:w_end + 1]
        event = delta_I[start:end + 1]

        event_mean = np.mean(event)

        # --- отрицательное событие ---
        if event_mean < 0:
            neg_peak = np.min(event)
            pos_peak = np.max(segment)

            if abs(pos_peak) < symmetry_ratio * abs(neg_peak):
                filtered_events.append((start, end))

        # --- положительное событие ---
        else:
            neg_peak = np.min(segment)
            pos_peak = np.max(event)

            if abs(neg_peak) < symmetry_ratio * abs(pos_peak):
                filtered_events.append((start, end))

    # --- удаление дублей ---
    events = list(set(filtered_events))
    events.sort()

    # ============================================
    # --- 4. НОВЫЙ БЛОК: УДАЛЕНИЕ КОРОТКИХ СОБЫТИЙ (< 2 мс) ---
    # ============================================
    min_duration_ms = 2  # Минимальная длительность в мс
    min_duration_points = int((min_duration_ms / 1000) / dt)  # Перевод в точки

    long_events = []
    for start, end in events:
        duration_points = end - start + 1
        if duration_points >= min_duration_points:
            long_events.append((start, end))

    events = long_events  # Обновляем список событий
    # ============================================

    # --- подсчёт ---
    negative_count = 0
    positive_count = 0

    for start, end in events:
        segment = delta_I[start:end + 1]
        if np.mean(segment) < 0:
            negative_count += 1
        else:
            positive_count += 1

    return events, negative_count, positive_count

def creating_table(filtered_events, delta_I, time, events_table, dt):
    for idx, (start, end) in enumerate(filtered_events, start=1):
        start_time = time[start]

        segment = delta_I[start:end + 1]
        min_val = np.min(segment)
        max_val = np.max(segment)
        amplitude = min_val if abs(min_val) >= abs(max_val) else max_val
        duration = (end - start + 1) * dt * 1000

        events_table.append((idx, start_time, duration, amplitude))

    events_table_sorted = sorted(
        events_table,
        key=lambda x: x[0],

    )
    return events_table_sorted


def count_events_by_sign(events, delta_I):
    """
    Подсчитывает количество положительных и отрицательных событий

    Parameters:
    filtered_events (list): Список событий в формате [(start, end), ...]
    delta_I (array): Массив значений сигнала

    Returns:
    tuple: (positive_count, negative_count)
    """
    positive_count = 0

    for start, end in events:
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)
        if event_mean > 0:  # положительное событие
            positive_count += 1

    negative_count = len(events) - positive_count
    return positive_count, negative_count


def calculation_one(values, a, k, positive_events, n_points, window, symmetry_ratio, dt, METOD, window_length, polyorder):

    if METOD == "EMA":
        delta_I = EMA_calculate_baseline(n_points, values, a)

    if METOD == "SG":
        m = np.zeros_like(values)
        m = savgol_filter(values, window_length, polyorder, mode="mirror")
        delta_I = values - m

    std_value, trigger_line, trigger = calculate_triggers(delta_I, k)

    print(f"для {METOD} СКО участка = {std_value:.6f}")
    print(f"ДЛЯ {METOD} Trigger level = {trigger_line:.6f}")

    if positive_events == 1:
        # Если галочка нажата - вызываем функцию для всех событий (сверху и снизу)
        raw_events = calculate_detecting_all(delta_I, trigger_line, trigger, n_points)
    else:
        # Если галочка не нажата - вызываем старую функцию только для нижних событий
        raw_events = calculate_detecting_down(delta_I, trigger_line, n_points)

        print(f"ДЛЯ {METOD} Найдено событий до фильтрации : {len(raw_events)}")
    filtered_events, negative_count, positive_count = filtering(raw_events, window, symmetry_ratio, n_points,
                                                                    delta_I,
                                                                    trigger_line, trigger, dt)

    if positive_events == 1:
        print(f"ДЛЯ {METOD} Событий после фильтрации: {len(filtered_events)}")
        print(f"ДЛЯ {METOD} Событий отрицательных: {negative_count}")
        print(f"ДЛЯ {METOD} Событий положительных: {positive_count}")
    else:
        print(f"ДЛЯ {METOD} Событий после фильтрации: {len(filtered_events)}")

    return std_value, trigger_line, trigger, raw_events, filtered_events, negative_count, positive_count, delta_I


def calculation_both(values, k, positive_events, n_points, window, symmetry_ratio, dt, METOD, values_ema, a, window_length, polyorder):
    if METOD == "SG и EMA":
        m = np.zeros_like(values)
        m = savgol_filter(values, window_length, polyorder, mode="mirror")
        delta_I = values - m

        ema_delta_I = EMA_calculate_baseline(n_points, values_ema, a)
        ema_std_value, ema_trigger_line, ema_trigger = calculate_triggers(ema_delta_I, k)

        print(f"для ЕМА СКО участка = {ema_std_value:.6f}")
        print(f"ДЛЯ ЕМА Trigger level = {ema_trigger_line:.6f}")

        if positive_events == 1:
            # Если галочка нажата - вызываем функцию для всех событий (сверху и снизу)
            ema_raw_events = calculate_detecting_all(ema_delta_I, ema_trigger_line, ema_trigger, n_points)
        else:
            # Если галочка не нажата - вызываем старую функцию только для нижних событий
            ema_raw_events = calculate_detecting_down(ema_delta_I, ema_trigger_line, n_points)

        print(f"ДЛЯ EMA Найдено событий до фильтрации : {len(ema_raw_events)}")
        ema_filtered_events, ema_negative_count, ema_positive_count = filtering(ema_raw_events, window, symmetry_ratio,
                                                                                n_points, ema_delta_I, ema_trigger_line,
                                                                                ema_trigger, dt)

        if positive_events == 1:
            print(f"ДЛЯ ЕМА Событий после фильтрации: {len(ema_filtered_events)}")
            print(f"ДЛЯ ЕМА Событий отрицательных: {ema_negative_count}")
            print(f"ДЛЯ ЕМА Событий положительных: {ema_positive_count}")
        else:
            print(f"ДЛЯ ЕМА Событий после фильтрации: {len(ema_filtered_events)}")


    std_value, trigger_line, trigger = calculate_triggers(delta_I, k)

    print(f"для SG СКО участка = {std_value:.6f}")
    print(f"ДЛЯ SG Trigger level = {trigger_line:.6f}")

    if positive_events == 1:
        # Если галочка нажата - вызываем функцию для всех событий (сверху и снизу)
        raw_events = calculate_detecting_all(delta_I, trigger_line, trigger, n_points)
    else:
        # Если галочка не нажата - вызываем старую функцию только для нижних событий
        raw_events = calculate_detecting_down(delta_I, trigger_line, n_points)

        print(f"ДЛЯ SG Найдено событий до фильтрации : {len(raw_events)}")
    filtered_events, negative_count, positive_count = filtering(raw_events, window, symmetry_ratio, n_points, delta_I,
                                                                trigger_line, trigger, dt)

    if positive_events == 1:
        print(f"ДЛЯ SG Событий после фильтрации: {len(filtered_events)}")
        print(f"ДЛЯ SG Событий отрицательных: {negative_count}")
        print(f"ДЛЯ SG Событий положительных: {positive_count}")
    else:
        print(f"ДЛЯ SG Событий после фильтрации: {len(filtered_events)}")


    return (std_value, trigger_line, trigger, raw_events,
            filtered_events, negative_count, positive_count,
            ema_std_value, ema_trigger_line, ema_trigger, ema_raw_events,
            ema_filtered_events, ema_negative_count, ema_positive_count, delta_I, ema_delta_I)

def save_summary_excel(filename, params_df, tables_dict):
    """
    tables_dict = {
        "События SG": SG_table,
        "События EMA": EMA_table
    }
    """

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        params_df.to_excel(writer, sheet_name='Параметры', index=False)

        for sheet_name, table in tables_dict.items():
            table.to_excel(writer, sheet_name=sheet_name, index=False)

def save_raw_events_excel(filename, params_df, data_dict, time, event_buffer, n_points):
    """
    data_dict = {
        "События SG": (events, signal),
        "События EMA": (events, signal)
    }
    """

    def create_event_table(events, signal):
        event_columns = {}

        for i, (start, end) in enumerate(events):
            seg_start = max(0, start - event_buffer)
            seg_end = min(n_points - 1, end + event_buffer)

            t_segment = time[seg_start:seg_end + 1]
            s_segment = signal[seg_start:seg_end + 1]

            event_columns[f"event_{i+1}_time"] = pd.Series(t_segment)
            event_columns[f"event_{i+1}_signal"] = pd.Series(s_segment)

        return pd.DataFrame(event_columns)

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        params_df.to_excel(writer, sheet_name='Параметры', index=False)

        for sheet_name, (events, signal) in data_dict.items():
            df = create_event_table(events, signal)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def save_full_signal_csv(filename, time, delta_I, ema_delta_I=None):
    """
    Сохраняет полный обработанный сигнал в CSV.
    Если ema_delta_I передан, создаются 2 файла: signal_SG.csv и signal_EMA.csv
    """
    import csv

    # Сохраняем SG сигнал (всегда)
    sg_filename = filename.replace('.csv', '_SG.csv')
    with open(sg_filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Time (s)', 'Delta_I_SG (pA)'])
        for t, d in zip(time, delta_I):
            writer.writerow([t, d])
    print(f"Сохранено: {sg_filename}")

    # Сохраняем EMA сигнал (если есть)
    if ema_delta_I is not None:
        ema_filename = filename.replace('.csv', '_EMA.csv')
        with open(ema_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Time (s)', 'Delta_I_EMA (pA)'])
            for t, d in zip(time, ema_delta_I):
                writer.writerow([t, d])
        print(f"Сохранено: {ema_filename}")