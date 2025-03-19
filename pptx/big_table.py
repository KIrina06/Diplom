### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт из файлов
from func import know_size, check_width, check_height, make_width, make_height, draw_free_raw, draw_special_head

### функции
@st.cache_resource
def draw_big_table(num_of_builders):
    fig = plt.figure(figsize = (15.2, 6.8))
    
    widths = [4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    heights = [5, 1, 1, 1, 1, 1, 1,
               1, 1, 1, 1, 1, 1, 1,
               1, 1, 1, 1, 1, 1, 1] # 21
    
    # создаем сетку графиков
    gs = GridSpec(ncols = 13, nrows = 21, figure = fig, wspace=0, hspace=0, width_ratios = widths, height_ratios = heights)
        
    # рисуем строки без данных из Excel, по сути шапку
    draw_special_head(fig, gs)
    
    
    ################################ рисуем строки с данными из Excel только серые #########################################
    # строки в т. ч. ИТР
    style = 2
    fontsize = 12
    color = ['w', '#A6A6A6', '#BFBFBF', '#606060', '#BFBFBF', '#606060']
    border_visibility = [False, True, False, True]
    
    arr = num_of_builders.iloc[1]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 4, data, fontsize, color, border_visibility, style)
    
    arr = num_of_builders.iloc[3]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 6, data, fontsize, color, border_visibility, style)
    
    arr = num_of_builders.iloc[6]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 10, data, fontsize, color, border_visibility, style)
    
    arr = num_of_builders.iloc[8]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 12, data, fontsize, color, border_visibility, style)
    
    arr = num_of_builders.iloc[11]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 16, data, fontsize, color, border_visibility, style)
    
    arr = num_of_builders.iloc[13]
    arr = arr.astype(str)
    data = np.insert(arr, 0, 'в т. ч. ИТР')
    draw_free_raw(fig, gs, 18, data, fontsize, color, border_visibility, style)
    ########################################################################################################################
    
    ########################################################################################################################
    rows = [3, 5, 9, 11]
    texts = ['ПЛАН', 'ФАКТ / ПРОГНОЗ', 'ПЛАН', 'ФАКТ / ПРОГНОЗ'] 
    colors = ['#BFBFBF', '#9DB5CB', '#767171', '#4F7393']
    for i in range(len(rows)):
        ax0 = plt.subplot(gs[rows[i], 0])
        plt.xlim(0, 1)
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
        
        plt.scatter(xmin + 0.05, (ymin + ymax) / 2, marker='s', facecolor=colors[i], s=100)
        
        ax0.text(xmin + 0.1, (ymin + ymax) / 2, texts[i], fontsize = fontsize, color = '#404040', va = 'center', ha = 'left', wrap = True)
        
        ax0.spines['left'].set_visible(False)
        ax0.spines['top'].set_color('#606060')
        ax0.spines['right'].set_visible(False)
        ax0.spines['bottom'].set_color('#606060')
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    ########################################################################################################################
    
    ########################################################################################################################
    rows = [15, 17]
    texts = ['ПЛАН', 'ФАКТ / ПРОГНОЗ'] 
    colors = ['#7F7F7F', '#4F7393']
    for i in range(len(rows)):
        ax0 = plt.subplot(gs[rows[i], 0])
        plt.xlim(0, 1)
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
        
        plt.scatter(xmin + 0.05, (ymin + ymax) / 2, marker='o', facecolor='None', edgecolor=colors[i], s=100, linewidths=2)
        
        ax0.text(xmin + 0.1, (ymin + ymax) / 2, texts[i], fontsize = fontsize, color = '#404040', va = 'center', ha = 'left', wrap = True)
        
        ax0.spines['left'].set_visible(False)
        ax0.spines['top'].set_color('#606060')
        ax0.spines['right'].set_visible(False)
        ax0.spines['bottom'].set_color('#606060')
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    ########################################################################################################################
        
    ########################################################################################################################
    rows = [7, 13, 19] 
    for i in range(len(rows)):
        ax0 = plt.subplot(gs[rows[i], 0])
        plt.xlim(0, 1)
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
        
        ax0.text(xmin, (ymin + ymax) / 2, 'ОТКЛ.', fontsize = fontsize, color = '#404040', va = 'center', ha = 'left', wrap = True)
        
        ax0.spines['left'].set_visible(False)
        ax0.spines['top'].set_color('#606060')
        ax0.spines['right'].set_visible(False)
        ax0.spines['bottom'].set_color('#606060')
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    ########################################################################################################################
    
    ########################################################################################################################
    rows = [3, 9, 15]
    ind = 0
    for i in range(0, len(num_of_builders) - 2, ++5):
        for j in range(len(num_of_builders.iloc[i])):
            ax0 = plt.subplot(gs[rows[ind], j + 1])
            plt.xlim(0, 1)
    
            text = num_of_builders.iloc[i, j]
            # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
            text, fontsize = make_width(fig, ax0, text, fontsize)
            text, fontsize = make_height(fig, ax0, text, fontsize)
        
            # ищем координаты размещения текста
            # (для шапки текст должен находится по центру графика)
            xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
            ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
            
            ax0.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#404040', va = 'center', ha = 'center', wrap = True)
            
            ax0.spines['left'].set_visible(False)
            ax0.spines['top'].set_color('#606060')
            ax0.spines['right'].set_visible(False)
            ax0.spines['bottom'].set_color('#606060')
            
            # убираем оси графика
            plt.xticks([])
            plt.yticks([])
    
        ind += 1
    ########################################################################################################################
    
    ############################### рисуем строки с данными из Excel не только серые #######################################
    rows = [5, 11, 17]
    ind = 0
    for i in range(2, len(num_of_builders) - 2, ++5):
        for j in range(len(num_of_builders.iloc[i])):
            ax0 = plt.subplot(gs[rows[ind], j + 1])
            plt.xlim(0, 1)
    
            text = num_of_builders.iloc[i, j]
            # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
            text, fontsize = make_width(fig, ax0, text, fontsize)
            text, fontsize = make_height(fig, ax0, text, fontsize)
        
            # ищем координаты размещения текста
            # (для шапки текст должен находится по центру графика)
            xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
            ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
    
            current_date = date.today()
            month = current_date.month - 1
            if j + 1 > month:
                color = '#404040'
            else:
                color = '#1F4E79'
            ax0.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'center', wrap = True)
            
            ax0.spines['left'].set_visible(False)
            ax0.spines['top'].set_color('#606060')
            ax0.spines['right'].set_visible(False)
            ax0.spines['bottom'].set_color('#606060')
            
            # убираем оси графика
            plt.xticks([])
            plt.yticks([])
    
        ind += 1
    
    ########################################################################################################################
    
    rows = [7, 13, 19]
    ind = 0
    for i in range(4, len(num_of_builders), ++5):
        for j in range(len(num_of_builders.iloc[i])):
            ax0 = plt.subplot(gs[rows[ind], j + 1])
            plt.xlim(0, 1)
    
            text = num_of_builders.iloc[i, j]
            # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
            text, fontsize = make_width(fig, ax0, text, fontsize)
            text, fontsize = make_height(fig, ax0, text, fontsize)
        
            # ищем координаты размещения текста
            # (для шапки текст должен находится по центру графика)
            xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
            ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
    
            if num_of_builders.iloc[i, j] >= 0:
                color = '#007a37'
            else:
                color = 'r'
            ax0.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'center', wrap = True)
            
            ax0.spines['left'].set_visible(False)
            ax0.spines['top'].set_color('#606060')
            ax0.spines['right'].set_visible(False)
            ax0.spines['bottom'].set_color('#606060')
            
            # убираем оси графика
            plt.xticks([])
            plt.yticks([])
    
        ind += 1
    
    ########################################################################################################################
    
    ############################################# рисуем строку с процентами ###############################################
    ax = plt.subplot(gs[20, 0])
    plt.xlim(0, 1)
    text = '%'
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
    
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_color('#606060')
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    ax.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#404040', va = 'center', ha = 'left', wrap = True)
    
    for i in range(len(num_of_builders.iloc[15])):
        ax0 = plt.subplot(gs[20, i + 1])
        plt.xlim(0, 1)
    
        text = str(num_of_builders.iloc[15, i]) + '%'
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax0, text, fontsize)
        text, fontsize = make_height(fig, ax0, text, fontsize)
        
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
    
        if num_of_builders.iloc[15, i] >= 100:
            color = '#007a37'
        else:
            color = 'r'
    
        ax0.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'center', wrap = True)
            
        ax0.spines['left'].set_visible(False)
        ax0.spines['top'].set_color('#606060')
        ax0.spines['right'].set_visible(False)
        ax0.spines['bottom'].set_visible(False)
            
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
    
    ########################################################################################################################
    
    ############################################## рисуем гистограммы ######################################################
    
    # рассчитываем высоту ячеек так, чтобы еще и график точечный поместился
    max_value_1 = max([num_of_builders.iloc[10, i] for i in range(12)])
    max_value_2 = max([num_of_builders.iloc[12, i] for i in range(12)])
    max_value = max([max_value_1, max_value_2])
    
    for i in range(1, 13):
        ax = plt.subplot(gs[0, i])
        plt.ylim(0, max_value * 1.3)
        plt.xlim(-0.5, 3.5)
        # названия столбцов
        labels = ['план дсо', 'факт/прогноз дсо', 'план стор', 'факт/прогноз стор']
    
        bar_cols = [num_of_builders.iloc[0, i - 1], num_of_builders.iloc[2, i - 1], num_of_builders.iloc[5, i - 1], num_of_builders.iloc[7, i - 1]]
        # отрисовка диаграммы
        bars = ax.bar(labels, bar_cols, facecolor = 'b', edgecolor = 'w', linewidth = 1.5)
    
        # раскрашиваем все
        bars[0].set_facecolor('#BFBFBF')
        bars[1].set_facecolor('#9DB5CB')
        bars[2].set_facecolor('#767171')
        bars[3].set_facecolor('#4F7393')
    
        # формируем границы диаграмм
        ax.spines['left'].set_color('#606060')
        ax.spines['left'].set_linestyle('dashed')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_color('#606060')
        ax.spines['right'].set_linestyle('dashed')
        ax.spines['bottom'].set_color('w')
        
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
    # достраиваем график
    ax1 = plt.subplot(gs[0, 1 : 13])
    
    # настраиваем оси
    plt.ylim(0, max_value * 1.3)
    plt.xlim(0, 4 * 12)
    
    plan = [num_of_builders.iloc[10, i] for i in range(12)]
    fact = [num_of_builders.iloc[12, i] for i in range(12)]
    x = [4 * i + 2 for i in range(12)]
    
    a = ax1.plot(x, plan)
    b = ax1.plot(x, fact)
    
    # оформляем график
    a[0].set_markersize(10)
    a[0].set_marker('o')
    a[0].set_markerfacecolor('None')
    a[0].set_markeredgecolor('#7F7F7F')
    a[0].set_color('#7F7F7F')
    
    b[0].set_markersize(10)
    b[0].set_marker('o')
    b[0].set_markerfacecolor('None')
    b[0].set_markeredgecolor('#4F7393')
    b[0].set_color('#4F7393')
    
    # подписи
    for i in range(len(plan)):
        plt.text(x[i], plan[i] + max_value // 10, plan[i], va = 'bottom', ha = 'center', color = '#7F7F7F', fontsize = 12)
        plt.text(x[i], fact[i] - max_value // 10, fact[i], va = 'top', ha = 'center', color = '#4F7393', fontsize = 12)
    
    # настраиваем график
    ax1.set_facecolor('None')
    ax1.spines['left'].set_visible(False)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['bottom'].set_visible(False)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    plt.savefig("system_photo/big_table.png", dpi=300, bbox_inches='tight')
    
