### импорт библиотек

import numpy as np
import matplotlib.pyplot as plt
import datetime
from datetime import date
import pandas as pd

### функции

# поиск минимальной даты
def min_date(dates):
    min_date = dates[0]
    for i in range(1, len(dates)):
        if min_date.year > dates[i].year:
            min_date = dates[i]
        elif min_date.year == dates[i].year:
            if min_date.month > dates[i].month:
                min_date = dates[i]
            elif min_date.month == dates[i].month:
                if min_date.day > dates[i].day:
                    min_date = dates[i]
    return pd.Timestamp(f'{min_date.year}-{min_date.month}-{min_date.day} 00:00:00')
    
# поиск максимальной даты
def max_date(dates):
    max_date = dates[0]
    for i in range(1, len(dates)):
        if max_date.year < dates[i].year:
            max_date = dates[i]
        elif max_date.year == dates[i].year:
            if max_date.month < dates[i].month:
                max_date = dates[i]
            elif max_date.month == dates[i].month:
                if max_date.day < dates[i].day:
                    max_date = dates[i]
    return pd.Timestamp(f'{max_date.year}-{max_date.month}-{max_date.day} 00:00:00')
    
# расчет координат
def count_coords(events, years_dict, sum_weights, year_marker, date_marker):
    dates = list()

    for i in range(len(events)):
        dates.append(events[i][date_marker])
    dates = [pd.Timestamp(f'{min_date(dates).year}-{1}-{1} 00:00:00')] + dates
    dates.append(pd.Timestamp(f'{max_date(dates).year}-{12}-{31} 00:00:00'))
    
    # превращаем даты в числа
    num_dates = list()
    date_min = dates[0].year % 100 + dates[0].month / 12 * 2 + dates[0].day / 300 * 2
    date_max = dates[len(dates) - 1].year % 100 + dates[len(dates) - 1].month / 12 * 2 + dates[len(dates) - 1].day / 300 * 2
    
    for i in range(len(dates)):
        num_dates.append(dates[i].year % 100 + dates[i].month / 12 * 2 + dates[i].day / 300 * 2)
    
    # масштабируем их
    for i in range(len(num_dates)):
        num_dates[i] = (num_dates[i] - date_min) / (date_max - date_min) * (years_dict[year_marker] / sum_weights * 100)
    
    num_dates.pop(0)
    num_dates.pop()

    return num_dates
    
# меняем цвет рамки графика
def change_color_frame(ax):
        ax.spines['bottom'].set_color('#595959')
        ax.spines['top'].set_color('#595959')
        ax.spines['left'].set_color('#595959')
        ax.spines['right'].set_color('#595959')
        return ax
    
# меняем цвет рамки шапки графика и делает заливку шапки
def change_color_frame_head(ax):
    ax.spines['bottom'].set_color('w')
    ax.spines['top'].set_color('w')
    ax.spines['left'].set_color('w')
    ax.spines['right'].set_color('w')
    ax.set_facecolor('#DAE8F2')
    return ax
    
# функция, которая узнает размер (ширину и высоту) элемента графика
def know_size(fig, x):
    r = fig.canvas.get_renderer()
    bb = x.get_window_extent(renderer=r)
    width = bb.width
    height = bb.height
    return width, height

# функция, проверяющая, не вылезает ли текст за пределы графика в ширину
# True - не вылезает
# False - вылезает
def check_width(fig, ax, text, fontsize):
    # добавляем ячейку с текстом на график
    r = fig.canvas.get_renderer()
    t = plt.text(0, 0, text, fontsize = fontsize, wrap=True)
    # узнаем размер графика
    width, height = know_size(fig, ax)
    # узнаем размер ячейки текста
    w, h = know_size(fig, t)
    # удаляем текст с графика (скорее скрываем, конечно)
    t.set_visible(False)
    # сравниваем ширину и выходим из функции
    if (w >= width):
        return False
    else:
        return True

# функция, проверяющая, не вылезает ли текст за пределы графика в высоту
# True - не вылезает
# False - вылезает
def check_height(fig, ax, text, fontsize):
    # добавляем ячейку с текстом на график
    r = fig.canvas.get_renderer()
    t = plt.text(0, 0, text, fontsize = fontsize, wrap=True)
    # узнаем размер графика
    width, height = know_size(fig, ax)
    # узнаем размер ячейки текста
    w, h = know_size(fig, t)
    # удаляем текст с графика (скорее скрываем, конечно)
    t.set_visible(False)
    # сравниваем высоту и выходим из функции
    if (h >= height):
        return False
    else:
        return True
        
# функция подбора ширины для текста
def make_width(fig, ax, text, fontsize):
    subtext = ''  # часть текста после предположительного переноса строки
    # пока текст не вмещается в нужную ширину, выполняем действия по его преобразованию
    while (not check_width(fig, ax, text, fontsize)):
        # ищем пробелы между словами, чтобы переносить целые слова на другую строку
        if (text.rfind(' ') != -1):
            # если нашли пробел, то разделяем текст на две части:
            # 1 - до пробела (оставляем в text)
            # 2 - после пробела (переносим эту часть в subtext)
            subtext = text[text.rfind(' ') + 1 : ] + ' ' + subtext
            text = text[0 : text.rfind(' ')]
            # проверяем, входит ли первая часть текста в наш график
            if check_width(fig, ax, text, fontsize):
                # если с первой частью все понятно, надо определяться с оставшейся
                # запускаем эту же функцию (так рекурсивно дойдем до конца текста)
                subtext, fontsize = make_width(fig, ax, subtext, fontsize)
                # обе части теперь входят в график по ширине, так что соединяем их и выходим
                text = text + '\n' + subtext
                return text, fontsize
        else:
            # если нет пробелов, то слово слишком длинное
            # единственный выход - это уменьшить шрифт, делаем это
            fontsize -= 0.5
    # весь наш текст теперь входит в график по ширине, но сделаем проверку,
    # чтобы исключить ситуации зацикливания и недообработки фрагмента subtext
    if subtext != '':
        subtext, fontsize = make_width(fig, ax, subtext, fontsize)
        text = text + '\n' + subtext
    return text, fontsize

# функция подбора высоты для текста
def make_height(fig, ax, text, fontsize):
    while (not check_height(fig, ax, text, fontsize)):
        # уменьшаем шрифт текста, пока текст не влезет в ячейку
        fontsize -= 0.5
        # text, fontsize = make_width(fig, ax, text, fontsize)
    return text, fontsize

# рисуем график-ячейку с текстом
def draw_graph_cell(fig, ax, text, fontsize, color, flag):
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
    
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
    # рисуем график-ячейку в зависимости от того, где в таблице эта ячейка находится
    if flag == 'head':
        # оформление
        ha = 'center'
        va = 'center'

        ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, ha = ha, va = va, wrap = True)
    elif flag == 'body':
        # оформление
        ha = 'left'
        va = 'center'
        
        ax.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, ha = ha, va = va, wrap = True)
    
    ax = change_color_frame(ax)
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])

    return ax

# отрисовка горизонтальной диаграммы
def make_horizontal_diagram(ax, forecasts, factplan, dates):
    # настраиваем оси диаграммы
    ax.set_xlim(0, 12)
    
    # названия столбцов
    labels = ['факт', 'план']
    
    # отрисовка горизонтальной столбчатой диаграммы
    bars = ax.barh(labels, forecasts, linestyle = '--', height = 0.65, facecolor = 'None', edgecolor = '#2E658E', linewidth = 1.5)
    ax.barh(labels, factplan, height = 0.7, facecolor = '#2972A7')
    
    for i in range(len(forecasts)):
        ax.text(forecasts[i], labels[i], '▲', fontsize = 30, ha = 'center', va = 'center', color = "w")
        ax.text(forecasts[i], labels[i], '▲', fontsize = 20, ha = 'center', va = 'center', color = "#2E658E")
        ax.text(forecasts[i], labels[i], dates[i], fontsize = 11, ha = 'right', va = 'bottom', color = "#003274", fontweight = 'bold')
    
    # изменяем оформление 1-го столбца диаграммы
    bars[1].set_facecolor('#B1CFE5')
    bars[1].set_edgecolor('None')
    
    # убираем оси диаграммы
    plt.xticks([])
    plt.yticks([])
    # раскрашиваем рамку диаграммы
    ax = change_color_frame(ax)

# функция отрисовки строки с диаграммой
def draw_row(fig, gs, row, key_event, sign, volume, act_completed, percent_completed, forecasts, factplan, dates, fontsize, color, flag):
    # рисуем график-ячейку с текстом key_event
    ax1 = draw_graph_cell(fig, plt.subplot(gs[row, 0]), key_event, fontsize, color, flag)

    # рисуем график-ячейку с текстом sign
    ax2 = draw_graph_cell(fig, plt.subplot(gs[row, 1]), sign, fontsize, color, flag)

    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, plt.subplot(gs[row, 2]), volume, fontsize, color, flag = 'head')

    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, plt.subplot(gs[row, 3]), act_completed, fontsize, color, flag = 'head')

    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, plt.subplot(gs[row, 4]), percent_completed, fontsize, color, flag = 'head')

    # отрисовка горизонтальной столбчатой диаграммы
    make_horizontal_diagram(plt.subplot(gs[row, 5:17]), forecasts, factplan, dates)

# функция отрисовки строки без диаграммы
def draw_row_without_diagram(fig, gs, i, mark, fontsize, flag, i0):
    ax1 = plt.subplot(gs[i + 1, 0:2])
    ax1.spines['right'].set_visible(False)
    text = f'НС, влияющие на ключевое событие {i // 2 + i0}: '
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax1, text, fontsize, '#1A3C7B', flag)

    ax2 = plt.subplot(gs[i + 1, 2:])
    ax2.spines['left'].set_visible(False)
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax2, mark, fontsize, '#00B050', flag)

# функция создания шапки таблицы
def draw_head(fig, gs, fontsize, color, flag, year):
    ################################# создаем шапку таблицы #####################################
    ax1 = plt.subplot(gs[0 : 2, 0])
    text = 'НАИМЕНОВАНИЕ КЛЮЧЕВОГО СОБЫТИЯ'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax1, text, fontsize, color, flag)
    ax1 = change_color_frame_head(ax1)
    #############################################################################################
    ax2 = plt.subplot(gs[0 : 2, 1])
    text = 'ПРИЗНАК'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax2, text, fontsize, color, flag)
    ax2 = change_color_frame_head(ax2)
    #############################################################################################
    ax3 = plt.subplot(gs[0, 2 : 5])
    text = 'ФИЗ. ОБЪЕМЫ\n(ГДЕ ВОЗМОЖНО)'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax3, text, fontsize, color, flag)
    ax3 = change_color_frame_head(ax3)
    #############################################################################################
    ax4 = plt.subplot(gs[0, 5 : 17])
    text = f'{year} ГОД'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax4, text, fontsize, color, flag)
    ax4 = change_color_frame_head(ax4)
    #############################################################################################
    ax5 = plt.subplot(gs[1, 2])
    text = 'Объем по ключевому событию'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax5, text, fontsize, color, flag)
    ax5 = change_color_frame_head(ax5)
    #############################################################################################
    ax6 = plt.subplot(gs[1, 3])
    text = 'Факт выпол.'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax6, text, fontsize, color, flag)
    ax6 = change_color_frame_head(ax6)
    #############################################################################################
    ax7 = plt.subplot(gs[1, 4])
    text = '% вып.'
    
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax7, text, fontsize, color, flag)
    ax7 = change_color_frame_head(ax7)
    #############################################################################################
    # рисуем ячейки с номерами с 1 по 12
    for i in range(12):
        text = i + 1
        ax = plt.subplot(gs[1, i + 5])
        draw_graph_cell(fig, ax, text, fontsize, color, flag)
        ax = change_color_frame_head(ax)
    ################################# шапка нарисована ##########################################

# функция отрисовки графика
def make_graph(gr, name, x, plan_y, fact_y, pred_y, month):
    current_month = month
    current_month_index = current_month - 1

    gr.plot(x, plan_y, '*-', color="#afabab", label="ПЛАН")
    gr.plot(x[:current_month_index], fact_y, 'o-', color="#1a3c7b", label="ФАКТ")
    gr.plot(x[11:], pred_y, 'D-', color="#4d98c4", label="ПРОГНОЗ")
    gr.set_title(name, loc="left", color="#1a3c7b")
    for i, txt in enumerate(plan_y):
        gr.text(x[i], plan_y[i], txt, color="#949292", verticalalignment='top')
    for i, txt in enumerate(fact_y):
        gr.text(x[i], fact_y[i], txt, color="#1a3c7b", verticalalignment='bottom')
    gr.text(x[11], pred_y, pred_y[0], color="#1a3c7b", verticalalignment='bottom')
    gr.axvline(current_month_index, color="#2452a6", linestyle='-.')
    gr.text(current_month_index-1, max(plan_y)*0.9, "ТЕКУЩАЯ ДАТА", color="#1a3c7b")

    gr.legend()

# функция отрисовки строки для численности персонала на строительной площадки
def draw_free_raw(fig, gs, row, data, fontsize, color, border_visibility, style):
    for i in range(13):
        ax = plt.subplot(gs[row, i])
        text = data[i]
        # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
        text, fontsize = make_width(fig, ax, text, fontsize)
        text, fontsize = make_height(fig, ax, text, fontsize)
            
        # ищем координаты размещения текста
        # (для шапки текст должен находится по центру графика)
        xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
        ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y

        
        if style == 1 and i == 0:     
            ax.text(xmin / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color[1], va = 'center', ha = 'left', wrap = True)
        elif style == 2 and i == 0:
            ax.text(xmax, (ymin + ymax) / 2, text, fontsize = fontsize, color = color[1], va = 'center', ha = 'right', wrap = True)
        else:
            ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color[1], va = 'center', ha = 'center', wrap = True)
        
            
        # оформление текста и ячейки в целом
        ax.set_facecolor(color[0])
        if i != 0:
            ax.spines['left'].set_color(color[2])
        ax.spines['top'].set_color(color[3])
        if i != 12:
            ax.spines['right'].set_color(color[4])
        ax.spines['bottom'].set_color(color[5])

        if i != 0:
            ax.spines['left'].set_visible(border_visibility[0])
        else:
            ax.spines['left'].set_visible(False)
        ax.spines['top'].set_visible(border_visibility[1])
        if i != 12:
            ax.spines['right'].set_visible(border_visibility[2])
        else:
            ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_visible(border_visibility[3])
        
        # убираем оси графика
        plt.xticks([])
        plt.yticks([])

# отрисовка специфической шапки для таблицы численности персонала на площадке
def draw_special_head(fig, gs):
    # цвета для строк
    # 0 - для фона
    # 1 - для текста
    # 2 - для границы левой
    # 3 - для границы верхней
    # 4 - для границы правой
    # 5 - для границы нижней
    
    # видимость границ ячейки
    # 0 - граница слева
    # 1 - граница сверху
    # 2 - граница справа
    # 3 - граница снизу
    
    fontsize = 14
    months = ['ПОКАЗАТЕЛЬ', 'ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    row = 1
    color = ['#DAE8F2', '#404040', '#BFBFBF', 'w', '#BFBFBF', 'w']
    border_visibility = [True, True, True, False]
    style = 1  # первые ячейки с оформлением текста слева 2 - справа
    draw_free_raw(fig, gs, row, months, fontsize, color, border_visibility, style)
    
    dso = ['ДСО (собственные силы)', '', '', '', '', '', '', '', '', '', '', '', '']
    row = 2
    color = ['#F2F2F2', '#404040', '#BFBFBF', 'w', '#BFBFBF', '#606060']
    border_visibility = [False, False, False, True]
    draw_free_raw(fig, gs, row, dso, fontsize, color, border_visibility, style)
    
    org = ['СТОРОННИЕ ОРГАНИЗАЦИИ', '', '', '', '', '', '', '', '', '', '', '', '']
    row = 8
    color = ['#F2F2F2', '#404040', '#BFBFBF', '#606060', '#BFBFBF', '#606060']
    border_visibility = [False, True, False, True]
    draw_free_raw(fig, gs, row, org, fontsize, color, border_visibility, style)
    
    allinall = ['ИТОГО', '', '', '', '', '', '', '', '', '', '', '', '']
    row = 14
    color = ['#F2F2F2', '#404040', '#BFBFBF', '#606060', '#BFBFBF', '#606060']
    border_visibility = [False, True, False, True]
    draw_free_raw(fig, gs, row, allinall, fontsize, color, border_visibility, style)
    
# надписи на круговой диаграмме в параметрах
def mini_func(pct):
    return "{:.0f}%".format(pct)

# функция отрисовки текста сверху по центру
def draw_text_top(fig, gs, row, col, text, fontsize, color):
    ax = plt.subplot(gs[row, col])
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
    
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
    ax.text((xmin + xmax) / 2, ymax, text, fontsize = 12, color = color, va = 'top', ha = 'center', wrap = True)
    
    # формируем границы диаграмм
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.set_facecolor('None')
    ax.patch.set_alpha(1)
    
    # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    
# отрисовка текста по центру
def draw_text_center(fig, gs, row, col, text, fontsize, color, bold, flag):
    ax = plt.subplot(gs[row, col])
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
            
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y

    if bold and flag == 1:
        ax.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, fontweight = 'bold', color = color, va = 'center', ha = 'left', wrap = True)
    elif bold and flag == 0:
        ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, fontweight = 'bold', color = color, va = 'center', ha = 'center', wrap = True)
    elif not bold and flag == 1:
        ax.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'left', wrap = True)
    else:
        ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'center', wrap = True)
   
    # убираем рамку
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.set_facecolor('None')
    ax.patch.set_alpha(1)
                
     # убираем оси графика
    plt.xticks([])
    plt.yticks([])

# отрисовка текста по центру (но мало параметров)
def draw_text(fig, gs, row, col, text, color):
    ax = plt.subplot(gs[row, col])
    fontsize = 14
    
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
        
    # ищем координаты размещения текста
    # (для шапки текст должен находится по центру графика)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
    
    ax.text((xmin + xmax) / 2, (ymin + ymax) / 2, text, fontsize = fontsize, color = color, va = 'center', ha = 'center', wrap = True)
    
    # убираем рамку
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
            
     # убираем оси графика
    plt.xticks([])
    plt.yticks([])