### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт файлов
from func import change_color_frame, change_color_frame_head, know_size, check_width, draw_head
from func import check_height, make_width, make_height, draw_graph_cell, make_horizontal_diagram, draw_row, draw_row_without_diagram

### функции
@st.cache_resource
def draw_key_events(general, ncols, nrows, i0, event_names, signs, volumes, fact_acts, proc_acts, fact_compl_dates, forecast_compl_dates, affects):
    x = 0.85 + len(event_names) * 1.4875
    fig = plt.figure(figsize = (15.2, x))

    # параметры "графика" и "подграфиков"
    flag = 'head'
    widths = [8, 3, 3, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    heights = [0] * nrows
    for i in range(nrows):
        if i == 0 or i == 1:
            heights[i] = 1
        else:
            if i % 2 == 0:
                heights[i] = 2.3
            else:
                heights[i] = 0.7
        
    # создаем сетку графиков
    gs = GridSpec(ncols = ncols, nrows = nrows, figure = fig, wspace=0, hspace=0, width_ratios = widths, height_ratios = heights)
    
    # параметры для текста по умолчанию
    fontsize = 14
    color = '#595959'
    
    # создаем шапку таблицы
    draw_head(fig, gs, fontsize, color, flag, general['date'][0].year)
    
    ########################### рисуем саму таблицу-график ######################################
    # параметры ячеек для таблицы (изменения)
    flag = 'body'
    
    # параметры текста для таблицы (изменения)
    
    # данные для строки
    ##################################################
    #   key_event - ключевое событие                 #
    #   sign - признак                               #
    #   volume - объем по ключевому событию          #
    #   act_completed - фактически выполненное       #
    #   percent_completed - процент выполненного     #
    ##################################################
    ind = 0
    for i in range(2, len(event_names) * 2 + 1, ++ 2):
        # получаем данные для строки
        key_event = event_names[ind]
        sign = signs[ind]
        volume = volumes[ind]
        act_completed = fact_acts[ind]
        percent_completed = proc_acts[ind]
        
        # данные для горизонтальной столбчатой диаграммы
        # обработка текущей даты
        now = general['date'][0].month + general['date'][0].day / 30 - 1
        # прогнозы (факт и план)
        forecasts = [forecast_compl_dates[ind].month - 1 + forecast_compl_dates[ind].day / 30, fact_compl_dates[ind].month - 1 + fact_compl_dates[ind].day / 30]
        # значения факт, план
        factplan = [now, 0]
        # значения дат
        forecast_date_month = f'{forecast_compl_dates[ind].month}'
        if forecast_compl_dates[ind].month // 10 == 0:
            forecast_date_month = f'0{forecast_compl_dates[ind].month}'
        
        forecast_date_day = f'{forecast_compl_dates[ind].day}'
        if forecast_compl_dates[ind].day // 10 == 0:
            forecast_date_day = f'0{forecast_compl_dates[ind].day}'

        fact_date_month = f'{fact_compl_dates[ind].month}'
        if fact_compl_dates[ind].month // 10 == 0:
            fact_date_month = f'0{fact_compl_dates[ind].month}'
        
        fact_date_day = f'{fact_compl_dates[ind].day}'
        if fact_compl_dates[ind].day // 10 == 0:
            fact_date_day = f'0{fact_compl_dates[ind].day}'
            
        dates = [f'{forecast_date_day}.{forecast_date_month}', f'{fact_date_day}.{fact_date_month}']
        
        # отрисовка строки с диаграммой
        draw_row(fig, gs, i, key_event, sign, volume, act_completed, percent_completed, forecasts, factplan, dates, fontsize, color, flag)
    
        # получаем данные для строки
        mark = affects[ind]
        
        # отрисовка строки без диаграммы
        draw_row_without_diagram(fig, gs, i, mark, fontsize, flag, i0)

        ind += 1
    
    # строим линию "текущее время"
    ax = plt.subplot(gs[1 :, 5 : ])
    # настраиваем график
    ax.set_facecolor('None')
    ax.patch.set_alpha(1)
    ax.spines['bottom'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    ax.set_xlim(0, 12)
    
    # получаем текущую дату
    current_date = general['date'][0]
    day = current_date.day
    month = current_date.month - 1
    
    # получаем коррдинаты для прямой
    x_date = month + day / 30
    ymin, ymax = ax.get_ylim()
    ax.plot([x_date, x_date], [ymin, ymax], linestyle = '--', color = '#00B050')
    ax.text(x_date, ymin, 'текущая дата', fontsize = 12, fontfamily = 'cursive', ha = 'center', va = 'top', color = '#00B050')
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
def draw_key_events_graph(general, key_events):
    # функция для подсчета количества столбцов и строк
    # для этого нужно вытащить список ключевых событий и посчитать их количество
    # если событий больше 4, то нужно переносить на другой слайд
    # циклом for для каждых 4 ключевых событий создаем слайд
    # посчитываем количество строк: 2 - шапка, по 2 - на каждое событие
    i0 = 0
    page = 0  # счетчик страниц
    for i in range(len(key_events['event_name'])):
        #if i == len(key_events['event_name']) - 1:
        if i % 4 == 3 or i == len(key_events['event_name']) - 1:
            # вставка графика
            nrows = 2 + 2 * (i - i0 + 1)  # 2 - на шапку таблицы, по 2 на каждое событие
            ncols = 17
            data = list(key_events['event_name'][i0 : i + 1])
            year = general['date'][0]
            draw_key_events(general, ncols, nrows, i0, list(key_events['event_name'][i0 : i + 1]), list(key_events['sign'][i0 : i + 1]), list(key_events['volume'][i0 : i + 1]), list(key_events['fact_act'][i0 : i + 1]), list(key_events['proc_act'][i0 : i + 1]), list(key_events['fact_compl_date'][i0 : i + 1]), list(key_events['forecast_compl_date'][i0 : i + 1]), list(key_events['affect'][i0 : i + 1]))
            i0 = i + 1
            plt.savefig(f"system_photo/key_events_{page}.png", dpi=300, bbox_inches='tight')
            page += 1