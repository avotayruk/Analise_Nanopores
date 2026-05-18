import numpy as np
from openpyxl import Workbook
from scipy.signal import savgol_filter
import pandas as pd
from scipy.ndimage import uniform_filter1d


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

def calculate_adaptive_triggers(delta_I, k, window_pts=50000):
    """
    Возвращает массивы trigger_line и trigger, меняющиеся в каждой точке.
    Оценивает локальное СКО через скользящее среднее квадратов.
    """
    local_mean_sq = uniform_filter1d(delta_I**2, size=window_pts, mode='nearest')
    local_std = np.sqrt(np.maximum(local_mean_sq, 1e-12))  # защита от отрицательных значений из-за float
    trigger_line = -k * local_std
    trigger = k * local_std
    return local_std, trigger_line, trigger



def filtering(raw_events, window, symmetry_ratio, n_points, delta_I, trigger_line, trigger, dt, weight_coeff=1.0, adaptive_trigger=False):
    min_distance_ms2 = 50
    min_distance_ms1 = 5
    min_distance_points2 = int((min_distance_ms2 / 1000) / dt)
    min_distance_points1 = int((min_distance_ms1 / 1000) / dt)

    raw_events_sorted = sorted(raw_events)

    # # --- 1. КЛАСТЕРИЗАЦИЯ ---
    # clusters = []
    # current_cluster = [raw_events_sorted[0]]
    # for i in range(1, len(raw_events_sorted)):
    #     start2, end2 = raw_events_sorted[i]
    #     start1, end1 = current_cluster[-1]
    #     distance = start2 - end1
    #     if min_distance_points1 < distance < min_distance_points2:
    #         current_cluster.append((start2, end2))
    #     else:
    #         clusters.append(current_cluster)
    #         current_cluster = [(start2, end2)]
    # clusters.append(current_cluster)
    # clean_raw_events = [cluster[0] for cluster in clusters if len(cluster) == 1]

    # --- 2. РАСШИРЕНИЕ ДО BASELINE ---
    expanded_events = []
    for start, end in raw_events:
        event_segment = delta_I[start:end + 1]
        event_mean = np.mean(event_segment)
        if event_mean < 0:
            i = start
            while i > 0 and delta_I[i] < 0: i -= 1
            new_start = i + 1
            i = end
            while i < n_points - 1 and delta_I[i] < 0: i += 1
            new_end = i - 1
        else:
            i = start
            while i > 0 and delta_I[i] > 0: i -= 1
            new_start = i + 1
            i = end
            while i < n_points - 1 and delta_I[i] > 0: i += 1
            new_end = i - 1
        expanded_events.append((new_start, new_end))

    # --- 3. ВЗВЕШЕННАЯ ПРОВЕРКА СИММЕТРИИ ---
    filtered_events = []
    for start, end in expanded_events:
        w_start = max(0, start - window)
        w_end = min(n_points - 1, end + window)

        segment = delta_I[w_start:w_end + 1]
        event = delta_I[start:end + 1]

        # 1. Локальная базовая линия (event_buffer)
        # Рекомендация: считать baseline без самой области события, чтобы не было смещения
        left_ctx = segment[:start - w_start]
        right_ctx = segment[end - w_start + 1:]
        context = np.concatenate([left_ctx, right_ctx])
        buffer_mean = np.mean(context) if len(context) > 0 else np.mean(segment)

        # Переводим в отклонения от локальной baseline
        dev_segment = segment - buffer_mean
        dev_event = event - buffer_mean
        event_dev_mean = np.mean(dev_event)

        # 2. Веса (без изменений, применяются к dev_segment)
        s_in = start - w_start
        e_in = end - w_start
        n_seg = len(dev_segment)  # <-- ИСПРАВЛЕНО: вынесено до if/else

        if weight_coeff > 1.0:
            center = (s_in + e_in) / 2.0
            dist = np.abs(np.arange(n_seg) - center)
            max_dist = max(center, n_seg - 1 - center)
            weights = 1.0 + (weight_coeff - 1.0) * (dist / max_dist) if max_dist > 0 else np.ones(n_seg)
        else:
            weights = np.ones(n_seg)

        # 3. Проверка противоположного пика относительно buffer_mean
        if event_dev_mean < 0:  # Событие отрицательное относительно локальной baseline
            main_peak = abs(np.min(dev_event))
            opp_mask = dev_segment > 0
            max_opp = np.max(dev_segment[opp_mask] * weights[opp_mask]) if np.any(opp_mask) else 0.0
            if max_opp < symmetry_ratio * main_peak:
                filtered_events.append((start, end))
        else:  # Событие положительное
            main_peak = np.max(dev_event)
            opp_mask = dev_segment < 0
            max_opp = np.max(np.abs(dev_segment[opp_mask]) * weights[opp_mask]) if np.any(opp_mask) else 0.0
            if max_opp < symmetry_ratio * main_peak:
                filtered_events.append((start, end))


    # --- Удаление дублей ---
    events = sorted(list(set(filtered_events)))

    # --- 4. МИНИМАЛЬНАЯ ДЛИТЕЛЬНОСТЬ ---
    min_duration_ms = 0.05
    min_duration_points = int((min_duration_ms / 1000) / dt)
    events = [(s, e) for s, e in events if (e - s + 1) >= min_duration_points]

    # --- Подсчет ---
    negative_count = sum(1 for s, e in events if np.mean(delta_I[s:e+1]) < 0)
    positive_count = len(events) - negative_count

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


def calculation_one(values, a, k, positive_events, n_points, window, symmetry_ratio, dt, METOD, window_length, polyorder, adaptive_trigger=False, weight_coeff=1.0):
    if METOD == "EMA":
        delta_I = EMA_calculate_baseline(n_points, values, a)
    else:
        m = savgol_filter(values, window_length, polyorder, mode="mirror")
        delta_I = values - m

    if adaptive_trigger:
        std_arr, trigger_line, trigger = calculate_adaptive_triggers(delta_I, k, window_pts=50000)
        global_std = np.std(delta_I)  # для вывода в консоль
        print(f"для {METOD} (адаптивный) Глобальное СКО = {global_std:.6f}")
    else:
        std_arr, trigger_line, trigger = calculate_triggers(delta_I, k)
        global_std = std_arr

    print(f"ДЛЯ {METOD} Trigger level (средний) = {np.mean(trigger_line) if adaptive_trigger else trigger_line:.6f}")

    if positive_events == 1:
        raw_events = calculate_detecting_all(delta_I, trigger_line, trigger, n_points)
    else:
        raw_events = calculate_detecting_down(delta_I, trigger_line, n_points)

    print(f"ДЛЯ {METOD} Найдено событий до фильтрации : {len(raw_events)}")

    filtered_events, negative_count, positive_count = filtering(
        raw_events, window, symmetry_ratio, n_points, delta_I, trigger_line, trigger, dt,
        weight_coeff=weight_coeff, adaptive_trigger=adaptive_trigger
    )

    print(f"ДЛЯ {METOD} Событий после фильтрации: {len(filtered_events)}")
    return global_std, trigger_line, trigger, raw_events, filtered_events, negative_count, positive_count, delta_I


def calculation_both(values, k, positive_events, n_points, window, symmetry_ratio, dt, METOD, values_ema, a, window_length, polyorder, adaptive_trigger=False, weight_coeff=1.0):
    # EMA часть
    ema_delta_I = EMA_calculate_baseline(n_points, values_ema, a)
    if adaptive_trigger:
        ema_std_arr, ema_trigger_line, ema_trigger = calculate_adaptive_triggers(ema_delta_I, k, window_pts=50000)
        print(f"для ЕМА (адаптивный) Глобальное СКО = {np.std(ema_delta_I):.6f}")
    else:
        ema_std_arr, ema_trigger_line, ema_trigger = calculate_triggers(ema_delta_I, k)

    ema_raw_events = calculate_detecting_all(ema_delta_I, ema_trigger_line, ema_trigger, n_points) if positive_events else calculate_detecting_down(ema_delta_I, ema_trigger_line, n_points)
    print(f"ДЛЯ EMA Найдено событий до фильтрации : {len(ema_raw_events)}")
    ema_filtered_events, ema_negative_count, ema_positive_count = filtering(
        ema_raw_events, window, symmetry_ratio, n_points, ema_delta_I, ema_trigger_line, ema_trigger, dt,
        weight_coeff=weight_coeff, adaptive_trigger=adaptive_trigger
    )

    # SG часть
    m = savgol_filter(values, window_length, polyorder, mode="mirror")
    delta_I = values - m
    if adaptive_trigger:
        std_arr, trigger_line, trigger = calculate_adaptive_triggers(delta_I, k, window_pts=50000)
        print(f"для SG (адаптивный) Глобальное СКО = {np.std(delta_I):.6f}")
    else:
        std_arr, trigger_line, trigger = calculate_triggers(delta_I, k)

    raw_events = calculate_detecting_all(delta_I, trigger_line, trigger, n_points) if positive_events else calculate_detecting_down(delta_I, trigger_line, n_points)
    print(f"ДЛЯ SG Найдено событий до фильтрации : {len(raw_events)}")
    filtered_events, negative_count, positive_count = filtering(
        raw_events, window, symmetry_ratio, n_points, delta_I, trigger_line, trigger, dt,
        weight_coeff=weight_coeff, adaptive_trigger=adaptive_trigger
    )

    return (np.std(delta_I), trigger_line, trigger, raw_events, filtered_events, negative_count, positive_count,
            np.std(ema_delta_I), ema_trigger_line, ema_trigger, ema_raw_events, ema_filtered_events, ema_negative_count, ema_positive_count, delta_I, ema_delta_I)


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