### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт из файлов
from func import know_size, check_width, check_height, make_width, make_height

### функции
@st.cache_resource
def draw_RD_SMR_graphs(general, RD_month, RD_accumulative):
    fig = plt.figure(figsize = (15.2, 6.8), frameon = False)
    
    widths = [2, 1, 1, 1, 1, 1, 1, 1, 1,
               1, 1, 1, 1, 0.5, 2, 1, 1, 1,
               1, 1, 1, 1, 1, 1, 1, 1, 1] # 27
    heights = [1, 1, 5, 1, 1, 1, 1, 1, 1]
    
    # создаем сетку графиков
    gs = GridSpec(ncols = 27, nrows = 9, figure = fig, wspace=0, hspace=0, width_ratios = widths, height_ratios = heights)
    
    fontsize = 14
    #################################### названия таблиц #########################################
    ax0 = plt.subplot(gs[0, 0 : 13])
    
    plt.xlim(0, 1)
    
    text = f'ВЫДАЧА КОРРЕКТИРОВКИ РД, НЕОБХОДИМОЙ ДЛЯ СМР В {general["date"][0].year} г., ед.'
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax0, text, fontsize)
    text, fontsize = make_height(fig, ax0, text, fontsize)
        
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
    
    ax0.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True)
    
    # убираем границы
    ax0.spines['left'].set_visible(False)
    ax0.spines['top'].set_visible(False)
    ax0.spines['right'].set_visible(False)
    ax0.spines['bottom'].set_visible(False)
            
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    ax1 = plt.subplot(gs[0, 14 : 27])
    
    plt.xlim(0, 1)
    
    text = f'ВЫДАЧА КОРРЕКТИРОВКИ РД, НЕОБХОДИМОЙ ДЛЯ СМР В {general["date"][0].year} г., ед.'
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax1, text, fontsize)
    text, fontsize = make_height(fig, ax1, text, fontsize)
        
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax1.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax1.get_ylim()  # получаем координаты начала и конца оси y
    
    ax1.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True)
    
    # убираем границы
    ax1.spines['left'].set_visible(False)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['bottom'].set_visible(False)
            
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    
    ax2 = plt.subplot(gs[1, 0 : 13])
    
    plt.xlim(0, 1)
    
    text = '(ПО МЕСЯЦАМ)'
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax2, text, fontsize)
    text, fontsize = make_height(fig, ax2, text, fontsize)
        
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax2.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax2.get_ylim()  # получаем координаты начала и конца оси y
    
    ax2.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True, fontweight = 'bold')
    
    # убираем границы
    ax2.spines['left'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['bottom'].set_visible(False)
            
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    
    ax2 = plt.subplot(gs[1, 14 : 27])
    
    plt.xlim(0, 1)
    
    text = '(НАКОПИТЕЛЬНО)'
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax2, text, fontsize)
    text, fontsize = make_height(fig, ax2, text, fontsize)
        
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax2.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax2.get_ylim()  # получаем координаты начала и конца оси y
    
    ax2.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True, fontweight = 'bold')
    
    # убираем границы
    ax2.spines['left'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['bottom'].set_visible(False)
            
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    ##############################################################################################
    
    ################################# разделительная полоса ######################################
    ax3 = plt.subplot(gs[:, 13])
    
    xmin, xmax = ax2.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax2.get_ylim()  # получаем координаты начала и конца оси y
    
    x = [(xmin + xmax) / 2, (xmin + xmax) / 2]
    ax3.plot(x, [ymin, ymax], color = '#A6A6A6')
    
    # убираем границы
    ax3.spines['left'].set_visible(False)
    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)
    ax3.spines['bottom'].set_visible(False)
            
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    ##############################################################################################
    
    #################################### полоски с месяцами ######################################
    months = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    for i in range(1, 13):
        ax = plt.subplot(gs[3, i])
        text = months[i - 1]
        
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax, text, fontsize)
        text, fontsize = make_height(fig, ax, text, fontsize)
        
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
        
        ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#595959', ha = 'center', va = 'center', wrap = True)
    
        # убираем границы
        ax.spines['left'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
    
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
    
    for i in range(15, 27):
        ax = plt.subplot(gs[3, i])
        text = months[i - 15]
        
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax, text, fontsize)
        text, fontsize = make_height(fig, ax, text, fontsize)
        
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
        
        ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#595959', ha = 'center', va = 'center', wrap = True)
    
        # убираем границы
        ax.spines['left'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
    
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
    ##############################################################################################
    
    ###################################### названия строк ########################################
    
    names_of_rows = ['ПЛАН', 'ФАКТ', 'ПРОГНОЗ', 'ОТКЛ.', '% ВЫП.']
    
    for i in range(len(names_of_rows)):
        ax1 = plt.subplot(gs[4 + i, 0])
        text = names_of_rows[i]
        
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax1, text, fontsize)
        text, fontsize = make_height(fig, ax1, text, fontsize)
        
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax1.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax1.get_ylim()  # получаем координаты начала и конца оси y
        
        ax1.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#595959', ha = 'left', va = 'center', wrap = True)
    
        # убираем границы
        ax1.spines['left'].set_visible(False)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.spines['bottom'].set_visible(False)
    
        # задаем цвет ячейки
        if i == 0 or i == 3:
            ax1.set_facecolor('#DAE8F2')
    
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
        ax2 = plt.subplot(gs[4 + i, 14])
        
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax2, text, fontsize)
        text, fontsize = make_height(fig, ax2, text, fontsize)
        
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax2.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax2.get_ylim()  # получаем координаты начала и конца оси y
        
        ax2.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#595959', ha = 'left', va = 'center', wrap = True)
    
        # убираем границы
        ax2.spines['left'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.spines['bottom'].set_visible(False)
    
        # задаем цвет ячейки
        if i == 0 or i == 3:
            ax2.set_facecolor('#DAE8F2')
    
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
        
    ##############################################################################################
    
    ##################################### заполняем таблицы ######################################
    
    for i in range(4, 9):
        for j in range(1, 13):
            ax = plt.subplot(gs[i, j])
            text = RD_month.iloc[i - 4, j - 1]
            if i == 8:
                text = str(text) + '%'
            
            # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
            text, fontsize = make_width(fig, ax, text, fontsize)
            text, fontsize = make_height(fig, ax, text, fontsize)
            
            # ищем координаты размещения текста
            # (для шапки текст должен находится по центру графика)
            xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
            ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
            if i == 5:
                color = '#1A3C7B'
            elif i == 7 and RD_month.iloc[i - 4, j - 1] <= 0:
                color = '#CD857F'
            elif i == 7 and RD_month.iloc[i - 4, j - 1] > 0:
                color = '#007A37'
            elif i == 8 and RD_month.iloc[i - 4, j - 1] < 100:
                color = '#CD857F'
            else:
                color = '#595959'
    
            ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, ha = 'center', va = 'center', wrap = True)
            
            # задаем цвет ячейки
            if i == 4 or i == 7:
                ax.set_facecolor('#DAE8F2')
            
            # убираем границы
            ax.spines['left'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_visible(False)
        
            # убираем оси графика
            plt.xticks([])
            plt.yticks([])
    
    
    for i in range(4, 9):
        for j in range(15, 27):
            ax = plt.subplot(gs[i, j])
            text = RD_accumulative.iloc[i - 4, j - 15]
            if i == 8:
                text = str(text) + '%'
            
            # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
            text, fontsize = make_width(fig, ax, text, fontsize)
            text, fontsize = make_height(fig, ax, text, fontsize)
            
            # ищем координаты размещения текста
            # (для шапки текст должен находится по центру графика)
            xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
            ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
            if i == 5:
                color = '#1A3C7B'
            elif i == 7 and RD_accumulative.iloc[i - 4, j - 15] <= 0:
                color = '#CD857F'
            elif i == 7 and RD_accumulative.iloc[i - 4, j - 15] > 0:
                color = '#007A37'
            elif i == 8 and RD_accumulative.iloc[i - 4, j - 15] < 100:
                color = '#CD857F'
            else:
                color = '#595959'
    
            ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, ha = 'center', va = 'center', wrap = True)
            
            # задаем цвет ячейки
            if i == 4 or i == 7:
                ax.set_facecolor('#DAE8F2')
            
            # убираем границы
            ax.spines['left'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_visible(False)
        
            # убираем оси графика
            plt.xticks([])
            plt.yticks([])
    
    ##############################################################################################
    
    #################################### рисуем гистограммы ######################################
    # рассчитываем высоту ячеек так, чтобы еще и график точечный поместился
    max_value_1 = max([RD_month.iloc[0, i] for i in range(12)])
    max_value_2 = max([RD_month.iloc[1, i] for i in range(12)])
    max_value_3 = max([RD_month.iloc[2, i] for i in range(12)])
    max_value = max([max_value_1, max_value_2, max_value_3])
    
    for i in range(1, 13):
        ax = plt.subplot(gs[2, i])
        plt.ylim(0, max_value * 1.1)
        plt.xlim(-0.5, 2.5)
        # названия столбцов
        labels = ['план', 'факт', 'прогноз']
    
        bar_cols = [RD_month.iloc[0, i - 1], RD_month.iloc[1, i - 1], RD_month.iloc[2, i - 1]]
        # отрисовка диаграммы
        bars = ax.bar(labels, bar_cols, facecolor = 'b', edgecolor = 'w', linewidth = 1.5)
        ax.text(labels[0], bar_cols[0], bar_cols[0], color = '#7F7F7F', fontsize = 9, ha = 'center', va = 'bottom', rotation = 60)
        ax.text(labels[1], bar_cols[1], bar_cols[1], color = '#1A3C7B', fontsize = 9, ha = 'center', va = 'bottom', rotation = 60)
        ax.text(labels[2], bar_cols[2], bar_cols[2], color = '#5B9BD5', fontsize = 9, ha = 'center', va = 'bottom', rotation = 60)
    
        # раскрашиваем все
        bars[0].set_facecolor('#7F7F7F')
        bars[1].set_facecolor('#1A3C7B')
        bars[2].set_facecolor('#5B9BD5')
        
        # формируем границы диаграмм
        ax.spines['left'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_color('#707070')
    
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])
    
    # попытка сделать легенду
    ax0 = plt.subplot(gs[2 : 4, 0])
    plt.ylim(0, 1)
    plt.xlim(0, 1)
        
    ax0.scatter(0.05, 0.3, marker = 's', facecolor = '#7F7F7F', s = 100)
    ax0.text(0.2, 0.3, '- ПЛАН', fontsize = 10, color = '#606060', va = 'center', ha = 'left', wrap = True)
    
    ax0.scatter(0.05, 0.6, marker = 's', facecolor = '#1A3C7B', s = 100)
    ax0.text(0.2, 0.6, '- ФАКТ', fontsize = 10, color = '#606060', va = 'center', ha = 'left', wrap = True)
    
    ax0.scatter(0.05, 0.9, marker = 's', facecolor = '#5B9BD5', s = 100)
    ax0.text(0.2, 0.9, '- ПРОГНОЗ', fontsize = 10, color = '#606060', va = 'center', ha = 'left', wrap = True)
    
    # формируем границы диаграмм
    ax0.spines['left'].set_visible(False)
    ax0.spines['top'].set_visible(False)
    ax0.spines['right'].set_visible(False)
    ax0.spines['bottom'].set_visible(False)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    ##############################################################################################
    
    ###################################### рисуем график #########################################
    max_value_1 = max([RD_accumulative.iloc[0, i] for i in range(12)])
    max_value_2 = max([RD_accumulative.iloc[1, i] for i in range(12)])
    max_value_3 = max([RD_accumulative.iloc[2, i] for i in range(12)])
    max_value = max([max_value_1, max_value_2, max_value_3])
    
    ax = plt.subplot(gs[2, 15 : 27])
    plt.xlim(0, 12)
    
    x = [i + 0.5 for i in range(12)]
    a = ax.plot(x, RD_accumulative.iloc[0], color = '#7F7F7F')
    general['date'][0].month
    
    x1 = [i + 0.5 for i in range(general['date'][0].month - 1)]
    b = ax.plot(x1, RD_accumulative.iloc[1, 0 : general['date'][0].month - 1], color = '#1A3C7B')
    
    x2 = [i + 0.5 for i in range(12)]
    c = ax.plot(x2, RD_accumulative.iloc[2, 0 : 12], color = '#1A3C7B', linestyle = 'dashed')
    
    # оформляем график
    a[0].set_markersize(10)
    a[0].set_marker('o')
    a[0].set_markerfacecolor('None')
    a[0].set_markeredgecolor('#7F7F7F')
    a[0].set_color('#7F7F7F')
    
    b[0].set_markersize(10)
    b[0].set_marker('o')
    b[0].set_markerfacecolor('None')
    b[0].set_markeredgecolor('#1A3C7B')
    b[0].set_color('#1A3C7B')
    
    c[0].set_markersize(10)
    c[0].set_marker('o')
    c[0].set_markerfacecolor('None')
    c[0].set_markeredgecolor('#1A3C7B')
    c[0].set_color('#1A3C7B')
    
    for i in range(len(x)):
        plt.text(x[i], RD_accumulative.iloc[0, i] + max_value // 20, RD_accumulative.iloc[0, i], va = 'bottom', ha = 'center', color = '#7F7F7F', fontsize = 10)
    
    for i in range(len(x1)):
        plt.text(x1[i], RD_accumulative.iloc[1, i] - max_value // 20, RD_accumulative.iloc[1, i], va = 'top', ha = 'center', color = '#1A3C7B', fontsize = 10)
    
    for i in range(len(x2)):
        plt.text(x2[i], RD_accumulative.iloc[2, i] - max_value // 20, RD_accumulative.iloc[2, i], va = 'top', ha = 'center', color = '#1A3C7B', fontsize = 10)
    
    
    # формируем границы диаграмм
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    # попытка сделать легенду
    ax0 = plt.subplot(gs[2 : 4, 14])
    plt.ylim(0, 1)
    plt.xlim(0, 1)
        
    ax0.scatter(0.1, 0.4, marker = 'o', facecolor = 'None', edgecolor = '#7F7F7F', s = 100, linewidths = 2)
    ax0.text(0.25, 0.4, '- ПЛАН', fontsize = 10, color = '#606060', va = 'center', ha = 'left', wrap = True)
    
    ax0.scatter(0.1, 0.7, marker = 'o', facecolor = 'None', edgecolor = '#1A3C7B', s = 100, linewidths = 2)
    ax0.text(0.25, 0.7, '- ФАКТ/\nПРОГНОЗ', fontsize = 10, color = '#606060', va = 'center', ha = 'left', wrap = True)
    
    # формируем границы диаграмм
    ax0.spines['left'].set_visible(False)
    ax0.spines['top'].set_visible(False)
    ax0.spines['right'].set_visible(False)
    ax0.spines['bottom'].set_visible(False)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
    plt.savefig("system_photo/RD_SMR_graphs.png", dpi=300, bbox_inches='tight')