#!/usr/bin/env python
# coding: utf-8

# # Импорт библиотек

# In[1]:


#!pip install pptx
#!pip install numpy
#!pip install datetime
#!pip install matplotlib
#!pip install pandas
#!pip install re
#!pip install xlrd
#!pip install psycopg2
#!pip install openpyxl


# In[2]:

import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
from datetime import datetime
import os

import pptx
from pptx import Presentation
from pptx.util import Cm
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE

import warnings
# # Импорт путей к картинкам

# In[3]:


icon_path = 'system_photo/icon.png'
icon2_path = 'system_photo/icon2.png'
background_path = 'system_photo/background.png'
plot_path = 'system_photo/plot.png'


# # Функции

# ## Функции, связанные непосредственно с python-pptx и презентацией

# ### Функции разметки

# In[4]:


def resize_text(rect_height, rect_width, text, fontsize):
    
    #text = "ИНВЕНТАРИЗАЦИЯ ПРОЕКТА СООРУЖЕНИЕ АЭС «РУППУР» НА ПРЕДМЕТ НАЛИЧИЯ РИСКОВ СРЫВА ПОnnnnnnnnnnnnnnnnСТАВОК ОБОРУДОВАНИЯ И КОМПЛЕКТУЮЩИХ ИЗ 3-Х СТРАН"
    fig = plt.figure(figsize = (rect_width, rect_height))
    
    gs = GridSpec(ncols = 1, nrows = 1, figure = fig, wspace=0, hspace=0)
    ax = plt.subplot(gs[: , :])
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
    
    #print(fontsize)
    xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
    ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y
            
    ax.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True)
            
     # убираем оси графика
    plt.xticks([])
    plt.yticks([])
    plt.show()
    #print(text)

    return text, fontsize


# In[5]:


# функция создания и отрисовки заголовка слайда
def name_of_slide(slide, name, subinf):
    left = Cm(3.23)
    top = Cm(-0.25)
    height = 1.16
    width = 16 - 1.27
    fontsize = 32
    name, fontsize = resize_text(height * 1.2, width * 1.5, name, fontsize)
    width = Inches(16 - 1.27)
    height = Cm(2.94)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    #tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = name
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.bold = True
    p.font.size = Pt(fontsize)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)

    if subinf != "":
        p1 = tf.add_paragraph()
        p1.text = "(" + subinf + ")"
        # Optionally, if you want to center the entire textbox including its vertical position
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        # Center the paragraph text
        p1.alignment = PP_ALIGN.LEFT
        p1.font.size = Pt(16)
        p1.font.name = "Arial Narrow"
        p1.font.color.rgb = RGBColor(32, 56, 100)


# In[6]:


def draw_table(shapes, top):
    rows = 10
    cols = 11
    left = Cm(0.41)
    #top = Cm(4.05)
    width = Cm(39.63)
    height = Cm(7.54)
    
    table = shapes.add_table(rows, cols, left, top, width, height).table
    
    cell = table.cell(0, 0)
    other_cell = table.cell(2, 0)
    cell.merge(other_cell)
    
    cell.text = "№"
    
    cell = table.cell(0, 1)
    other_cell = table.cell(2, 1)
    cell.merge(other_cell)
    
    cell.text = "ПОКАЗАТЕЛЬ"
    
    cell = table.cell(0, 2)
    other_cell = table.cell(2, 2)
    cell.merge(other_cell)
    
    cell.text = "ВЫПОЛНЕНИЕ ПЛАНА ПЕРИОДА ПО КОНТРАКТАЦИИ, %"
    
    cell = table.cell(0, 3)
    other_cell = table.cell(0, 10)
    cell.merge(other_cell)
    
    cell.text = "ОСВОЕНИЕ (стоимость выполненных и принятых работ на проекте)"
    
    cell = table.cell(1, 3)
    other_cell = table.cell(2, 3)
    cell.merge(other_cell)
    
    cell.text = "План года*"
    
    table.cell(1, 4).text = "План периода"
    table.cell(2, 4).text = "(ЯНВ-АПР)"
    
    cell = table.cell(1, 5)
    other_cell = table.cell(1, 6)
    cell.merge(other_cell)
    
    cell.text = "Факт за период"
    
    cell = table.cell(2, 5)
    other_cell = table.cell(2, 6)
    cell.merge(other_cell)
    
    cell.text = "(ЯНВ-АПР)**"
    
    cell = table.cell(1, 7)
    other_cell = table.cell(2, 8)
    cell.merge(other_cell)
    
    cell.text = "% вып. от плана периода"
    
    cell = table.cell(1, 9)
    other_cell = table.cell(2, 9)
    cell.merge(other_cell)
    
    cell.text = "% вып. от плана года"
    
    table.cell(1, 10).text = "Прогноз % вып."
    table.cell(2, 10).text = "на конец года"
    
    table.cell(3, 0).text = "1"
    table.cell(4, 0).text = "1.1"
    table.cell(5, 0).text = "1.2"
    table.cell(6, 0).text = "1.3"
    table.cell(7, 0).text = "1.4"
    table.cell(8, 0).text = "2"
    
    cell = table.cell(9, 0)
    other_cell = table.cell(9, 1)
    cell.merge(other_cell)
    
    cell.text = "ИТОГО"
    
    table.cell(3, 1).text = "КАПИТАЛЬНЫЕ СТРОИТЕЛЬНЫЕ ЗАТРАТЫ"
    table.cell(4, 1).text = "СМР"
    table.cell(5, 1).text = "ОБОРУДОВАНИЕ"
    table.cell(6, 1).text = "ПИР"
    table.cell(7, 1).text = "ПРОЧИЕ"
    table.cell(8, 1).text = "АССОЦИИРОВАННЫЕ КАПИТАЛЬНЫЕ ЗАТРАТЫ"

    cnt = 0
    # set column widths
    for col in table.columns:
        if cnt == 1:
            col.width = Cm(9)
        elif cnt == 0:
            col.width = Cm(1.5)
        elif cnt ==5 or cnt == 6:
            col.width = Cm(2)
        cnt += 1
    
    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
            
            fill = cell.fill
            fill.solid()
            if cnt <= 2:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1

    return table


# In[7]:


def make_graph(gr, name, x, plan_y, fact_y, pred_y):
    current_month = datetime.now().month
    current_month_index = current_month - 1

    gr.plot(x, plan_y, '*-', color="#afabab", label="ПЛАН")
    gr.plot(x[:4], fact_y, 'o-', color="#1a3c7b", label="ФАКТ")
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


# In[ ]:





# ## Функции, связанные с внутренним построением объектом (без python-pptx)

# ### Функции для графиков

# In[8]:


# надписи на круговой диаграмме в параметрах
def func(pct):
    return "{:.0f}%".format(pct)


# In[9]:


# перевод даты в число
def to_integer(dt_time):
    return dt_time.year % 100 * 100 / 30 + dt_time.month * 3 / 13 + dt_time.day / 31


# In[10]:


# масштабирование данных [0, 100]
def scaling_number(data):
    return (data - data. min()) / (data.max() - data.min()) * 96 + 2


# In[11]:


# перевод даты в числа
def date_to_nums(date):
    data = [int] * len(date)
    
    for i in range(len(date)):
        data[i] = to_integer(date[i])
    return data


# In[12]:


# какое-то шаманство с датами-числами (Не дай Бог это редактировать...)
def shamanstvo_s_datami(date, year):
    x = date_to_nums(date)
    x = np.array(x)
    
    old_date = []
    ind = []
    for i in range(len(x)):
        if x[i] < (year) % 100 * 100 / 30 + 1 / 13 + 1 / 31:
            old_date.append(x[i])
            ind.append(i)
    for i in range(len(old_date) - 1, -1, -1):
        old_date[i] = ((year) % 100 * 100 / 30 + 1 / 13 + 1 / 31) - (len(old_date) - i) * 2
        #old_date[i] = x[ind[len(ind)-1]] - (len(old_date) - i) * 2
    
    for i in range(len(ind)):
        x[ind[i]] = old_date[i]
    return x


# In[13]:


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
    return min_date.day, min_date.month, min_date.year


# In[14]:


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
    return max_date.day, max_date.month, max_date.year


# ### Функции декорирования (графики/текст/...)

# In[15]:


# меняем цвет рамки графика
def change_color_frame(ax):
    ax.spines['bottom'].set_color('#595959')
    ax.spines['top'].set_color('#595959')
    ax.spines['left'].set_color('#595959')
    ax.spines['right'].set_color('#595959')
    return ax


# In[16]:


# меняем цвет рамки шапки графика и делает заливку шапки
def change_color_frame_head(ax):
    ax.spines['bottom'].set_color('w')
    ax.spines['top'].set_color('w')
    ax.spines['left'].set_color('w')
    ax.spines['right'].set_color('w')
    ax.set_facecolor('#DAE8F2')
    return ax


# ### Функции для нормальной вставки текста в график

# In[17]:


# функция, которая узнает размер (ширину и высоту) элемента графика
def know_size(fig, x):
    r = fig.canvas.get_renderer()
    bb = x.get_window_extent(renderer=r)
    width = bb.width
    height = bb.height
    return width, height


# In[18]:


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


# In[19]:


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


# In[20]:


# функция подбора ширины для текста
def make_width(fig, ax, text, fontsize):
    subtext = ''  # часть текста после предположительного переноса строки
    # пока текст не вмещается в нужную ширину, выполняем действия по его преобразованию
    while (not check_width(fig, ax, text, fontsize)):
        # ищем пробелы между словами, чтобы переносить целые слова на другую строку
        if (str(text).rfind(' ') != -1):
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


# In[21]:


# функция подбора высоты для текста
def make_height(fig, ax, text, fontsize):
    while (not check_height(fig, ax, text, fontsize)):
        # уменьшаем шрифт текста, пока текст не влезет в ячейку
        fontsize -= 0.5
        # text, fontsize = make_width(fig, ax, text, fontsize)
    return text, fontsize


# ### Функции отрисовки

# In[22]:


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


# In[23]:


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


# In[24]:


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


# In[25]:


# функция создания шапки таблицы
def draw_head(fig, gs, fontsize, color, flag):
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


# In[26]:


# функция отрисовки строки без диаграммы
def draw_row_without_diagram(fig, gs, i, mark, fontsize, flag):
    ax1 = plt.subplot(gs[i + 1, 0:2])
    ax1.spines['right'].set_visible(False)
    text = f'НС, влияющие на ключевое событие {i // 2}: '
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax1, text, fontsize, '#1A3C7B', flag)

    ax2 = plt.subplot(gs[i + 1, 2:])
    ax2.spines['left'].set_visible(False)
    # рисуем график-ячейку с текстом
    draw_graph_cell(fig, ax2, mark, fontsize, '#00B050', flag)


# In[27]:


# функция отрисовки графика
def make_graph(gr, name, x, plan_y, fact_y, pred_y):
    current_month = datetime.now().month
    current_month_index = current_month - 1

    gr.plot(x, plan_y, '*-', color="#afabab", label="ПЛАН")
    gr.plot(x[:4], fact_y, 'o-', color="#1a3c7b", label="ФАКТ")
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


# In[28]:


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


# In[29]:


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


# In[30]:


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


# In[31]:


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


# In[32]:


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

st.title('Презентация')
# # Считывание данных из Excel-формы

# In[33]:
uploaded_file = st.file_uploader("Выберите файл", type=['xlsx'])

if uploaded_file is not None:
    st.write("Файл успешно загружен!")

#file = 'form.xlsx'
file = uploaded_file
# In[34]:


xl = pd.ExcelFile(file)


# In[35]:


#xl.sheet_names


# In[36]:


general = xl.parse('general')


# In[37]:


#general


# In[38]:


graph_1_lvl = xl.parse('graph_1_lvl')


# In[39]:


#graph_1_lvl


# In[40]:


key_events = xl.parse('key_events')


# In[41]:


#key_events


# In[42]:


#key_events['event_name'][:]


# In[43]:


program_execution = xl.parse('program_execution')


# In[44]:


#program_execution


# In[45]:


num_of_builders = xl.parse('num_of_builders')


# In[46]:


#num_of_builders


# In[47]:


values = [num_of_builders.iloc[2][col] - num_of_builders.iloc[0][col] for col in num_of_builders.columns]

num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 4, values, axis= 0 ))
num_of_builders2.columns = num_of_builders.columns
#num_of_builders2


# In[48]:


values = [num_of_builders2.iloc[7][col] - num_of_builders2.iloc[5][col] for col in num_of_builders2.columns]

num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 9, values, axis= 0 ))
num_of_builders.columns = num_of_builders2.columns
#num_of_builders


# In[49]:


values = [num_of_builders.iloc[0][col] + num_of_builders.iloc[5][col] for col in num_of_builders.columns]

num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 10, values, axis= 0 ))
num_of_builders2.columns = num_of_builders.columns
#num_of_builders2


# In[50]:


values = [num_of_builders2.iloc[1][col] + num_of_builders2.iloc[6][col] for col in num_of_builders2.columns]

num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 11, values, axis= 0 ))
num_of_builders.columns = num_of_builders2.columns
#num_of_builders


# In[51]:


# num_of_builders.drop(12, inplace=True)


# In[52]:


values = [num_of_builders.iloc[2][col] + num_of_builders.iloc[7][col] for col in num_of_builders.columns]

num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 12, values, axis= 0 ))
num_of_builders2.columns = num_of_builders.columns
#num_of_builders2


# In[53]:


values = [num_of_builders2.iloc[3][col] + num_of_builders2.iloc[8][col] for col in num_of_builders2.columns]

num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 13, values, axis= 0 ))
num_of_builders.columns = num_of_builders2.columns
#num_of_builders


# In[54]:


values = [num_of_builders.iloc[12][col] - num_of_builders.iloc[10][col] for col in num_of_builders.columns]

num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 14, values, axis= 0 ))
num_of_builders2.columns = num_of_builders.columns
#num_of_builders2


# In[55]:


values = [round((num_of_builders2.iloc[12][col] / num_of_builders2.iloc[10][col] * 100), 0) for col in num_of_builders2.columns]

num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 15, values, axis= 0 ))
num_of_builders.columns = num_of_builders2.columns
#num_of_builders


# In[56]:


RD_month = xl.parse('RD_month')


# In[57]:


#RD_month


# In[58]:


values = [RD_month.iloc[1][col] - RD_month.iloc[0][col] for col in RD_month.columns]

RD_month2 = pd.DataFrame(np.insert(RD_month.values, 3, values, axis= 0 ))
RD_month2.columns = RD_month.columns
#RD_month2


# In[59]:


months = ['jan', 'feb',	'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
# получаем текущую дату
month = general['date'][0].month
for i in range(len(RD_month2.iloc[3])):
    if months.index(RD_month2.columns[i]) >= month - 1:
        RD_month2.iloc[3, i] = 0
#RD_month2


# In[60]:


values = []
for col in RD_month2.columns:
    if RD_month2.iloc[0][col] != 0:
        values.append(round((RD_month2.iloc[1][col] / RD_month2.iloc[0][col] * 100), 0))
    else:
        values.append(0)

RD_month = pd.DataFrame(np.insert(RD_month2.values, 4, values, axis= 0 ))
RD_month.columns = RD_month2.columns
#RD_month


# In[61]:


RD_accumulative = xl.parse('RD_accumulative')


# In[62]:


#RD_accumulative


# In[63]:


values = [RD_accumulative.iloc[1][col] - RD_accumulative.iloc[0][col] for col in RD_accumulative.columns]

RD_accumulative2 = pd.DataFrame(np.insert(RD_accumulative.values, 3, values, axis= 0 ))
RD_accumulative2.columns = RD_accumulative.columns
#RD_accumulative2


# In[64]:


months = ['jan', 'feb',	'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
# получаем текущую дату
month = general['date'][0].month
for i in range(len(RD_accumulative2.iloc[3])):
    if months.index(RD_accumulative2.columns[i]) >= month - 1:
        RD_accumulative2.iloc[3, i] = 0
#RD_accumulative2


# In[65]:


values = []
for col in RD_accumulative2.columns:
    if RD_accumulative2.iloc[0][col] != 0:
        values.append(round((RD_accumulative2.iloc[1][col] / RD_accumulative2.iloc[0][col] * 100), 0))
    else:
        values.append(0)

RD_accumulative = pd.DataFrame(np.insert(RD_accumulative2.values, 4, values, axis= 0 ))
RD_accumulative.columns = RD_accumulative2.columns
#RD_accumulative


# In[66]:


inventory = xl.parse('inventory')


# In[67]:


#inventory


# In[68]:


xl.close()


# Filter out all warnings
warnings.filterwarnings("ignore")


# ### Титульный слайд
st.subheader('Титульный лист')
# In[69]:
text = st.text_input("Название сооружения: ")
if text != "":
    general['object_name'][0] = text

types = ["сооружение", "корабль"]
text = st.selectbox("Тип сооружения: ", types)
if text != "":
    general['object_type'][0] = text

date_str = st.text_input("Дата (для рассмотрения на Операционном комитете) (ДД.ММ.ГГГГ):")
if date_str != "":
    # --- Преобразуем строку в объект datetime ---
    try:
        selected_date = datetime.strptime(date_str, "%d.%m.%Y").date()
        general['date'][0] = pd.Timestamp(selected_date)
    except ValueError:
        st.error("Неверный формат даты. Используйте ДД.ММ.ГГГГ.")

date_str = st.text_input("Текущая дата:")

if date_str != "":
    # --- Преобразуем строку в объект datetime ---
    try:
        selected_date = datetime.strptime(date_str, "%d.%m.%Y").date()
        general['date_now'][0] = pd.Timestamp(selected_date)
    except ValueError:
        st.error("Неверный формат даты. Используйте ДД.ММ.ГГГГ.")

text = st.text_input("ФИО докладчика: ")
if text != "":
    general['fio'][0] = text

text = st.text_input("Должность докладчика: ")
if text != "":
    general['job'][0] = text


object_name = general['object_name'][0]


# In[70]:


month = general['date'][0].month
if month // 10 == 0: 
    month = f'0{month}'

day = general['date'][0].day
if day // 10 == 0: 
    day = f'0{day}'


# ## График 1-ого уровня

# In[71]:
st.subheader('График 1-ого уровня')

year = date.today().year


# In[72]:


fig = plt.figure(figsize=(20,10))

# Разметка
plt.xlim(0, 100)
plt.ylim(0, 4)

# Скрытие осей
ax = plt.gca()
plt.axis('off')

# подготавливаем данные для графика
events = graph_1_lvl['event_name']  # названия ключевых событий 
date_plan = graph_1_lvl['date_plan']  # даты, утвержденные графиком
date_fact_forecast = graph_1_lvl['date_fact_forecast']  # даты по факту/прогнозу

# рисуем сами временные линии
plt.plot([0, 98.2], [1, 1], [0, 98.2], [2, 2], [0, 98.2], [3.8, 3.8], color="#BDD7EE", linewidth=30, zorder = 1)
plt.arrow(x=0, y=1, dx=99.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)
plt.arrow(x=0, y=2, dx=99.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)
plt.arrow(x=0, y=3.8, dx=99.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)

# начинаем наносить точки
# чтобы нормально нанести дату на график, 
# мы будем переводить ее в число и масштабировать его в отрезке [0, 100]

# сначала нанесем даты, утвержденные графиком (y = 2)
# переводим даты в числа и масштабируем их

min_d1 = [] * 3
min_d2 = [] * 3
min_d1 = min_date(date_plan)
min_d2 = min_date(date_fact_forecast)
min_d = [] * 3

if min_d1[2] < min_d2[2]:
    min_d = min_d1
elif min_d1[2] > min_d2[2]:
    min_d = min_d2
else:
    if min_d1[1] < min_d2[1]:
        min_d = min_d1
    elif min_d1[1] > min_d2[1]:
        min_d = min_d2
    else:
        if min_d1[0] < min_d2[0]:
            min_d = min_d1
        else:
            min_d = min_d2

max_d1 = [] * 3
max_d2 = [] * 3
max_d1 = max_date(date_plan)
max_d2 = max_date(date_fact_forecast)
max_d = [] * 3

if max_d1[2] > max_d2[2]:
    max_d = max_d1
elif max_d1[2] < max_d2[2]:
    max_d = max_d2
else:
    if max_d1[1] > max_d2[1]:
        max_d = max_d1
    elif max_d1[1] < max_d2[1]:
        max_d = max_d2
    else:
        if max_d1[0] > max_d2[0]:
            max_d = max_d1
        else:
            max_d = max_d2

min_d = list(min_d)
max_d = list(max_d)

if min_d[1] == 1:
    min_d[2] -= 1
    min_d[1] = 12
else:
    min_d[1] -= 1

if max_d[1] == 12:
    max_d[2] += 1
    max_d[1] = 1
else:
    max_d[1] += 1

t1 = pd.Timestamp(f'{min_d[2]}-{min_d[1]}-{min_d[0]} 00:00:00')
t2 = pd.Timestamp(f'{max_d[2]}-{max_d[1]}-{max_d[0]} 00:00:00')

x_date = []
x_date.append(t1)
for i in range(len(date_plan)):
    x_date.append(date_plan[i])
x_date.append(t2)

x = shamanstvo_s_datami(x_date, year = date.today().year)
x_plan = np.array(x)
x_plan = scaling_number(x_plan)

# наносим
for i in range(1, len(x_plan) - 1):
    plt.scatter(x_plan[i], 2, marker='o', facecolor='none', edgecolor='#2E75B6', s=800, linewidths=7, zorder = 2)


# нанесем даты по факту/прогнозу (y = 1)
# переводим даты в числа и масштабируем их
x_date = []
x_date.append(t1)
for i in range(len(date_fact_forecast)):
    x_date.append(date_fact_forecast[i])
x_date.append(t2)

x = shamanstvo_s_datami(x_date, year = date.today().year)
x_fact_forecast = np.array(x)
x_fact_forecast = scaling_number(x_fact_forecast)

# наносим
for i in range(1, len(x_fact_forecast) - 1):
    if date_fact_forecast[i - 1] > general['date'][0]:
        plt.scatter(x_fact_forecast[i], 1, marker='o', facecolor='none', linestyle=':', edgecolor='#1F435F', s=800, linewidths=7, zorder = 2)
    else:
        if date_fact_forecast[i - 1] > date_plan[i - 1]:
            plt.scatter(x_fact_forecast[i], 1, marker='o', facecolor='none', edgecolor='red', s=800, linewidths=7, zorder = 2)
        else:
            plt.scatter(x_fact_forecast[i], 1, marker='o', facecolor='none', edgecolor='#00B050', s=800, linewidths=7, zorder = 2)


# Подписи
# для значений факт/прогноз
for i in range(1, len(x_fact_forecast) - 1):
    month = date_fact_forecast[i - 1].month
    if month // 10 == 0: 
        month = f'0{month}'
    
    day = date_fact_forecast[i - 1].day
    if day // 10 == 0: 
        day = f'0{day}'
    if i % 3 == 0:
        plt.text(x_fact_forecast[i], 0.7, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
    elif i % 3 == 1:
        plt.text(x_fact_forecast[i], 0.5, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
    else:
        plt.text(x_fact_forecast[i], 0.3, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))

# для значений утвержденных планом
for i in range(1, len(x_plan) - 1):
    month = date_plan[i - 1].month
    if month // 10 == 0: 
        month = f'0{month}'
    
    day = date_plan[i - 1].day
    if day // 10 == 0: 
        day = f'0{day}'
    if i % 3 == 0:
        plt.text(x_plan[i], 1.7, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
    elif i % 3 == 1:
        plt.text(x_plan[i], 1.5, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
    else:
        plt.text(x_plan[i], 1.3, f'{day}.{month}.{date_plan[i - 1].year}', va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))


for i, txt in enumerate(events):
    if i % 5 == 0:
        plt.text(x_plan[i + 1], 3.5, events[i], va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
        plt.annotate("", xy=(x_fact_forecast[i + 1], 1), xytext=(x_plan[i + 1], 2), arrowprops=dict(arrowstyle="->"), color = '#003274', zorder = 4)
    elif i % 5 == 1:
        plt.text(x_plan[i + 1], 3.2, events[i], va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
        plt.annotate("", xy=(x_fact_forecast[i + 1], 1), xytext=(x_plan[i + 1], 2), arrowprops=dict(arrowstyle="->"), color = '#003274', zorder = 4)
    elif i % 5 == 2:
        plt.text(x_plan[i + 1], 2.9, events[i], va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
        plt.annotate("", xy=(x_fact_forecast[i + 1], 1), xytext=(x_plan[i + 1], 2), arrowprops=dict(arrowstyle="->"), color = '#003274', zorder = 4)
    elif i % 5 == 3:
        plt.text(x_plan[i + 1], 2.6, events[i], va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
        plt.annotate("", xy=(x_fact_forecast[i + 1], 1), xytext=(x_plan[i + 1], 2), arrowprops=dict(arrowstyle="->"), color = '#003274', zorder = 4)
    else:
        plt.text(x_plan[i + 1], 2.3, events[i], va = 'center', ha = 'center', color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
        plt.annotate("", xy=(x_fact_forecast[i + 1], 1), xytext=(x_plan[i + 1], 2), arrowprops=dict(arrowstyle="->"), color = '#003274', zorder = 4)
        
# Добавление флажка
x_flag = x_plan[len(x_plan)-2]
y_flag = 2.2
plt.plot(x_flag, y_flag, marker='|', markersize=60, markerfacecolor='#2E75B6', markeredgecolor='#2E75B6', zorder = 5)
plt.plot(x_flag + 1, y_flag + 0.19, marker='>', markersize=25, markerfacecolor='#2E75B6', markeredgecolor='#2E75B6', zorder = 5)
plt.plot(x_flag + 1, y_flag + 0.12, marker='>', markersize=25, markerfacecolor='#2E75B6', markeredgecolor='none', zorder = 5)

# Подпись годов
x1 = 0
x2 = 100
mark_year = 0
min_year = 0
max_year = 0
# НАДО ОСТАЛЬНЫЕ СЛУЧАИ РАССМОТРЕТЬ!!!!!!!!!
for i in range(len(x_date) - 2, 0, -1):
    if mark_year == 0:
        mark_year = x_date[i].year
    else:
        if mark_year != x_date[i].year and mark_year >= year:
            x1 = (x_plan[i] + x_plan[i + 1]) / 2
            plt.text((x1 + x2)/2, 3.8, f'{mark_year}', ha='center', color='#003274', fontsize=14)
            plt.plot(x1, 3.8, marker='|', markersize=30, markerfacecolor='w', markeredgecolor='w', zorder = 3)
            x2 = x1
            mark_year = x_date[i].year
        elif mark_year < year and max_year == 0:
            max_year = mark_year
        elif max_year != 0 and min_year == 0 and i == 1 and x_date[i].year != max_year:
            x1 = 0
            min_year = x_date[i].year
            plt.text((x1 + x2)/2, 3.8, f'{min_year} - {max_year}', ha='center', color='#003274', fontsize=14)

plt.savefig("system_photo/1lvlgraph.png", dpi=300, bbox_inches='tight')
#plt.show()
#st.pyplot(fig)
st.image('system_photo/1lvlgraph.png')


# ## Ключевые события

# In[73]:
st.subheader('Ключевые события')

def draw_key_events(ncols, nrows, event_names, signs, volumes, fact_acts, proc_acts, fact_compl_dates, forecast_compl_dates, affects):
    fig = plt.figure(figsize = (15.2, 6.8))

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
    draw_head(fig, gs, fontsize, color, flag)
    
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
        draw_row_without_diagram(fig, gs, i, mark, fontsize, flag)

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

    plt.savefig(f"system_photo/keyevents.png", dpi=300, bbox_inches='tight')
    #plt.show()
    st.image('system_photo/keyevents.png')

# In[74]:


# функция для подсчета количества столбцов и строк
# для этого нужно вытащить список ключевых событий и посчитать их количество
# если событий больше 4, то нужно переносить на другой слайд
# циклом for для каждых 4 ключевых событий создаем третий слайд
# посчитываем количество строк: 2 - шапка, по 2 - на каждое событие
i0 = 0
for i in range(len(key_events['event_name'])):
    if i % 4 == 3 : #or i == len(key_events['event_name']) - 1:  # после каждого 4-ого события создаем слайд с ключевыми событиями  
        #slide = ppt.slides.add_slide(blank_slide_layout)
        #shapes = slide.shapes
        # добавление названия слайда
        #subinf = "С УКАЗАНИЕМ ФИЗИЧЕСКИХ ОБЪЕМОВ РАБОТ"
        #name_of_slide(slide, f'КЛЮЧЕВЫЕ СОБЫТИЯ {general["date"][0].year} ГОДА', subinf)
        # вставка графика
        nrows = 2 + 2 * (i + 1)  # 2 - на шапку таблицы, по 2 на каждое событие
        ncols = 17
        draw_key_events(ncols, nrows, key_events['event_name'][i0 : i + 1], key_events['sign'][i0 : i + 1], key_events['volume'][i0 : i + 1], key_events['fact_act'][i0 : i + 1], key_events['proc_act'][i0 : i + 1], key_events['fact_compl_date'][i0 : i + 1], key_events['forecast_compl_date'][i0 : i + 1], key_events['affect'][i0 : i + 1])
        i0 = i + 1
        # на слайд
        #pic = slide.shapes.add_picture(plot_path, Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))


# ## Выполнение плана по освоению

# In[75]:
st.subheader('Выполнение плана по освоению')

f, ax = plt.subplots(1, 1)
f.set_size_inches(11.46, 3.85)
plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
plan_y = np.array([8, 21, 44, 101, 167, 298, 358, 447, 547, 648, 825, 1205])
fact_y = np.array([8, 21, 44, 117])
pred_y = np.array([1205])
make_graph(ax, "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (НАКОПИТЕЛЬНО), млн. долл.", x, plan_y, fact_y, pred_y)

plt.savefig("system_photo/plotplancompl.png", dpi=300, bbox_inches='tight')
#plt.show()
st.image('system_photo/plotplancompl.png')

# In[76]:


plan_dso = [9977, 11427]
plan_s = [7848, 6792]
x = ['ПЛАН', 'ФАКТ']

fig = plt.figure(figsize=(2.9, 2.19))
ax = fig.add_subplot()

ax.spines['top'].set_visible(False)
ax.spines['left'].set_visible(False)
ax.spines['right'].set_visible(False)

ax.set_yticklabels([])
ax.tick_params(axis='y', length=0)

ax.bar(x, plan_dso, color="#397073", label="ДСО (собственные силы)")
ax.bar(x, plan_s, bottom=plan_dso, color="#bfbfbf", label="Сторонние подрядчики")

for i, v in enumerate(plan_dso):
    ax.text(i, v/2, f"{v}", ha="center", va="center", fontsize=14, color="#bfbfbf", fontweight="bold")

for i, v in enumerate(plan_s):
    ax.text(i, v/2 + plan_dso[i], f"{v}", ha="center", va="center", fontsize=14, color="#397073", fontweight="bold")

for i, v in enumerate(plan_s):
    ax.text(i, v + plan_dso[i], f"{v + plan_dso[i]}", ha="center", va="bottom", fontsize=11, color="#595959", fontweight="bold")

plt.title(f"ЧИСЛЕННОСТЬ СТРОИТЕЛЬНОГО ПЕРСОНАЛА \nНА ПЛОЩАДКЕ, чел", fontsize=14, color="#1a3c7b", fontweight="semibold")
plt.legend(bbox_to_anchor=(-1.05, 1), loc='upper left')
plt.savefig("system_photo/bar1.png", dpi=300, bbox_inches='tight')
#plt.show()
st.image('system_photo/bar1.png')


# ## Выполнение плана по освоению по структуре затрат
st.subheader('Выполнение плана по освоению по структуре затрат')
# In[77]:


f, ax = plt.subplots(2, 2)
f.set_size_inches(15.2, 7.4)
plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
plan_y = np.array([0, 0, 0, 0, 1, 2, 2, 3, 6, 6, 7, 12])
fact_y = np.array([0, 0, 0, 0])
pred_y = np.array([12])
make_graph(ax[0, 0], "ПИР", x, plan_y, fact_y, pred_y)

plan_y = np.array([0, 3, 11, 24, 43, 118, 130, 156, 190, 218, 289, 342])
fact_y = np.array([0, 3, 11, 27])
pred_y = np.array([342])
make_graph(ax[0, 1], "ОБОРУДОВАНИЕ", x, plan_y, fact_y, pred_y)

plan_y = np.array([0, 0, 6, 40, 78, 124, 164, 219, 267, 317, 397, 646])
fact_y = np.array([0, 0, 6, 53])
pred_y = np.array([646])
make_graph(ax[1, 0], "СМР", x, plan_y, fact_y, pred_y)

plan_y = np.array([0, 0, 0, 0, 0, 0, 0, 0, 6, 6, 6, 12])
fact_y = np.array([0, 0, 0, 0])
pred_y = np.array([12])
make_graph(ax[1, 1], "ПРОЧЕЕ", x, plan_y, fact_y, pred_y)

plt.savefig("system_photo/4plot.png", dpi=300, bbox_inches='tight')
#plt.show()
#st.pyplot(fig)
st.image('system_photo/4plot.png')


# ## Выполнение показателя освоение

# ## Выполнение плана по реализации

# In[78]:
st.subheader('Выполнение плана по реализации')

f, ax = plt.subplots(1, 1)
f.set_size_inches(11.46, 3.85)
plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
plan_y = np.array([8, 21, 44, 101, 167, 298, 358, 447, 547, 648, 825, 1205])
fact_y = np.array([8, 21, 44, 117])
pred_y = np.array([1205])
make_graph(ax, "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (НАКОПИТЕЛЬНО), млн. долл.", x, plan_y, fact_y, pred_y)

plt.savefig("system_photo/plot2.png", dpi=300, bbox_inches='tight')
#plt.show()
st.image('system_photo/plot2.png')


# ## Приложения

# ## Статус выдачи рд на объем смр

# In[79]:
st.subheader('Статус выдачи рд на объем смр')

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

plt.savefig("system_photo/2bigplot.png", dpi=300, bbox_inches='tight')
#plt.show()
#st.pyplot(fig)
st.image('system_photo/2bigplot.png')


# ## Численность строительного персонала на площадке

# In[80]:
st.subheader('Численность строительного персонала на площадке')

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

plt.savefig("system_photo/bigtable.png", dpi=300, bbox_inches='tight')
# plt.show()
#st.pyplot(fig)
st.image('system_photo/bigtable.png')


# ## Инвентаризация проекта на предмет наличия рисков срыва поставок оборудования и комплектующих из 3-х стран

# In[81]:
st.subheader('Инвентаризация проекта на предмет наличия рисков срыва поставок оборудования и комплектующих из 3-х стран')

fig = plt.figure(figsize = (15.2, 6.8))

widths = [4.5, 1, 1, 1, 5, 1, 4, 1, 1]
heights = [1, 1, 1, 1, 1.5, 1] # 21

# создаем сетку графиков
gs = GridSpec(ncols = 9, nrows = 6, figure = fig, wspace=0, hspace=0, width_ratios = widths, height_ratios = heights)

all_mln = inventory['contract_no_risk'][0] + inventory['contract_risk'][0] + inventory['no_contract_no_risk'][0] + inventory['no_contract_risk'][0]

ax0 = plt.subplot(gs[0, :])
plt.xlim(0, 1)

text = f'Общий объем предстоящего к поставке оборудования до конца реализации проекта составляет {all_mln} млн долл., из них:'
fontsize = 14

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text, fontsize = make_width(fig, ax0, text, fontsize)
text, fontsize = make_height(fig, ax0, text, fontsize)
    
# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax0.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax0.get_ylim()  # получаем координаты начала и конца оси y
        
ax0.text(xmin, (ymin + ymax) / 2, text, fontsize = fontsize, color = '#1A3C7B', va = 'center', ha = 'left', wrap = True)

# убираем рамку
ax0.spines['left'].set_visible(False)
ax0.spines['top'].set_visible(False)
ax0.spines['right'].set_visible(False)
ax0.spines['bottom'].set_visible(False)
        
 # убираем оси графика
plt.xticks([])
plt.yticks([])

# подсчет данных для слайда
contract = [inventory['contract_no_risk'][0], inventory['contract_risk'][0]]
proc_contract = [round(inventory['contract_no_risk'][0] / all_mln * 100, 0), round(inventory['contract_risk'][0] / all_mln * 100, 0)]

no_contract = [inventory['no_contract_no_risk'][0], inventory['no_contract_risk'][0]]
proc_no_contract = [round(inventory['no_contract_no_risk'][0] / all_mln * 100, 0), round(inventory['no_contract_risk'][0] / all_mln * 100, 0)]

itogo = [contract[1] + no_contract[1], proc_contract[1] + proc_no_contract[1]]

# рисуем саму круговую диаграмму
ax1 = plt.subplot(gs[2 :, 4])

inner = [proc_no_contract[1], proc_no_contract[0], proc_contract[1], proc_contract[0]]
outer = [proc_no_contract[0] + proc_no_contract[1], proc_contract[0] + proc_contract[1]]

colors_inner = ['#9CB2C8', '#6184A7', '#ADC7B6', '#6A9A79']
colors_outer = ['#6184A7', '#6A9A79']

wp = {'linewidth': 1, 'edgecolor': "white", 'width': 0.6}

# внутреннее кольцо
wedges, texts, autotexts = ax1.pie(inner,
                                  radius=1.5,
                                  autopct=lambda pct: func(pct),
                                  colors=colors_inner,
                                  startangle=90,
                                  wedgeprops=wp,
                                  textprops=dict(color = "white", fontweight = "bold", fontsize = 14),
                                  explode=[0.02, 0.02, 0.02, 0.02]) 

# изменяем границы для 2-го и 4-го сегмента
wedgeprops = [
    {'linewidth': 2, 'edgecolor': "#FFA500", 'linestyle': 'dashed', 'width': 0.6}, 
    {'linewidth': 2, 'edgecolor': "white", 'width': 0.6},
    {'linewidth': 2, 'edgecolor': "#FFA500", 'linestyle': 'dashed', 'width': 0.6}, 
    {'linewidth': 2, 'edgecolor': "white", 'width': 0.6},
]
for i, wedge in enumerate(wedges):
    wedge.set_linewidth(wedgeprops[i]['linewidth'])
    wedge.set_edgecolor(wedgeprops[i]['edgecolor'])
    if 'linestyle' in wedgeprops[i]: 
        wedge.set_linestyle(wedgeprops[i]['linestyle']) 

# внешнее кольцо
wedges_outer, texts_outer = ax1.pie(outer, radius=1.65, startangle=90, colors=colors_outer, wedgeprops=dict(width=0.09, edgecolor="white", linewidth=6))

# надписи на сегментах внутреннего круга
for i, autotext in enumerate(autotexts):
    # получаем центр сегмента
    x, y = autotext.get_position()
    
    # увеличиваем радиус для надписи
    radius = 0.3
    
    # пересчитываем координаты надписи, чтобы она была ближе к краю
    angle = (wedges[i].theta1 + wedges[i].theta2) / 2 
    x_pos = x + radius * np.cos(np.deg2rad(angle))
    y_pos = y + radius * np.sin(np.deg2rad(angle))

    # устанавливаем новые координаты для надписи
    autotext.set_position((x_pos, y_pos))

# надпись в центре
text = all_mln
fontsize = 50

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text, fontsize = make_width(fig, ax1, text, fontsize)
text, fontsize = make_height(fig, ax1, text, fontsize)
    
# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax1.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax1.get_ylim()  # получаем координаты начала и конца оси y

ax1.text((xmin + xmax) / 2, (ymin + ymax) / 2 - ymax / 10, text, fontsize = fontsize, color = '#1A3C7B', va = 'bottom', ha = 'center', wrap = True)

text2 = 'млн долл.'
fontsize = 28

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text2, fontsize = make_width(fig, ax1, text2, fontsize)
text2, fontsize = make_height(fig, ax1, text2, fontsize)

ax1.text((xmin + xmax) / 2, (ymin + ymax) / 2 - ymax / 10, text2, fontsize = fontsize, color = '#1A3C7B', va = 'top', ha = 'center', wrap = True)

# убираем рамку
ax1.spines['left'].set_visible(False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.spines['bottom'].set_visible(False)
        
 # убираем оси графика
plt.xticks([])
plt.yticks([])

draw_text(fig, gs, 1, 1, 'млн долл.', '#404040')
draw_text(fig, gs, 1, 2, '%', '#404040')
draw_text(fig, gs, 1, 7, 'млн долл.', '#404040')
draw_text(fig, gs, 1, 8, '%', '#404040')

draw_text_center(fig, gs, 2, 0, 'НЕ ЗАКОНТРАКТОВАНО, в т.ч.', fontsize = 11, color = '#45607B', bold = True, flag = 0)
draw_text_center(fig, gs, 2, 6, 'ЗАКОНТРАКТОВАНО, в т.ч.', fontsize = 11, color = '#5A8467', bold = True, flag = 0)
draw_text_center(fig, gs, 2, 1, no_contract[0] + no_contract[1], fontsize = 12, color = '#45607B', bold = True, flag = 1)
draw_text_center(fig, gs, 2, 2, str(proc_no_contract[0] + proc_no_contract[1]) + '%', fontsize = 12, color = '#45607B', bold = True, flag = 1)
draw_text_center(fig, gs, 2, 7, contract[0] + contract[1], fontsize = 12, color = '#5A8467', bold = True, flag = 1)
draw_text_center(fig, gs, 2, 8, str(proc_contract[0] + proc_contract[1]) + '%', fontsize = 12, color = '#5A8467', bold = True, flag = 1)
draw_text_center(fig, gs, 3, 1, no_contract[0], fontsize = 12, color = '#404040', bold = False, flag = 0)
draw_text_center(fig, gs, 3, 2, str(proc_no_contract[0]) + '%', fontsize = 12, color = '#404040', bold = False, flag = 0)
draw_text_center(fig, gs, 3, 7, contract[0], fontsize = 12, color = '#404040', bold = False, flag = 0)
draw_text_center(fig, gs, 3, 8, str(proc_contract[0]) + '%', fontsize = 12, color = '#404040', bold = False, flag = 0)


ax = plt.subplot(gs[3, 0])
fontsize = 12

plt.ylim(0, 1)
plt.xlim(0, 1)

# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y

ax.scatter(xmin + 0.05, (ymin + ymax) / 2, marker = 's', facecolor = '#45607B', s = 100)
ax.text(xmin + 0.1, (ymin + ymax) / 2, 'Риск отсутствует', fontsize = 12, color = '#404040', va = 'center', ha = 'left', wrap = True)

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


ax = plt.subplot(gs[4, 0])
fontsize = 12
text = '     Риск срыва поставок ввиду санкционного давления на РФ (использование импортных комплектующих - идет подбор альтернативных)'

plt.ylim(0, 1)
plt.xlim(0, 1)

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text, fontsize = make_width(fig, ax, text, fontsize)
text, fontsize = make_height(fig, ax, text, fontsize)

# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y


ax.scatter(xmin + 0.05, ymax - 0.075, marker = 's', facecolor = '#9CB2C8', edgecolor = '#D84612', linestyle = '--', s = 100, linewidths = 1)
ax.text(xmin, ymax, text, fontsize = 12, color = '#D84612', va = 'top', ha = 'left', wrap = True)

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


ax = plt.subplot(gs[3, 6])
fontsize = 12

plt.ylim(0, 1)
plt.xlim(0, 1)

# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y

ax.scatter(xmin + 0.05, (ymin + ymax) / 2, marker = 's', facecolor = '#6A9A79', s = 100)
ax.text(xmin + 0.1, (ymin + ymax) / 2, 'Риск отсутствует', fontsize = 12, color = '#404040', va = 'center', ha = 'left', wrap = True)

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


ax = plt.subplot(gs[4, 6])
fontsize = 12
text = '     Риск срыва поставок ввиду санкционного давления на РФ'

plt.ylim(0, 1)
plt.xlim(0, 1)

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text, fontsize = make_width(fig, ax, text, fontsize)
text, fontsize = make_height(fig, ax, text, fontsize)

# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y


ax.scatter(xmin + 0.05, ymax - 0.075, marker = 's', facecolor = '#ADC7B6', edgecolor = '#D84612', linestyle = '--', s = 100, linewidths = 1)
ax.text(xmin, ymax, text, fontsize = 12, color = '#D84612', va = 'top', ha = 'left', wrap = True)

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

draw_text_top(fig, gs, 4, 1, str(no_contract[1]), fontsize = 12, color = '#D84612')
draw_text_top(fig, gs, 4, 2, str(proc_no_contract[1]) + '%', fontsize = 12, color = '#D84612')
draw_text_top(fig, gs, 4, 7, contract[1], fontsize = 12, color = '#D84612')
draw_text_top(fig, gs, 4, 8, str(proc_contract[1]) + '%', fontsize = 12, color = '#D84612')

ax = plt.subplot(gs[5, 6 : 9])
fontsize = 12
text = f'Итого риск поставок по законтрактованному и не законтрактованному оборудованию составляет {no_contract[1] + contract[1]} млн долл. или {proc_no_contract[1] + proc_contract[1]}%'

plt.ylim(0, 1)
plt.xlim(0, 1)

# изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
text, fontsize = make_width(fig, ax, text, fontsize)
text, fontsize = make_height(fig, ax, text, fontsize)

# ищем координаты размещения текста
# (для шапки текст должен находится по центру графика)
xmin, xmax = ax.get_xlim()  # получаем координаты начала и конца оси x
ymin, ymax = ax.get_ylim()  # получаем координаты начала и конца оси y


ax.text(xmin, ymax, text, fontsize = 12, color = '#D84612', va = 'top', ha = 'left', wrap = True)

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

ax4 = plt.subplot(gs[2, 0 : 4])

plt.ylim(0, 1)
plt.xlim(0, 1)

ax4.plot([0, 0.95], [0.2, 0.2], color = '#45607B')
ax4.scatter(0.95, 0.2, marker = 'o', color = '#45607B')

# формируем границы диаграмм
ax4.spines['left'].set_visible(False)
ax4.spines['top'].set_visible(False)
ax4.spines['right'].set_visible(False)
ax4.spines['bottom'].set_visible(False)
ax4.set_facecolor('None')
ax4.patch.set_alpha(1)

# убираем оси графика
plt.xticks([])
plt.yticks([])

ax5 = plt.subplot(gs[2, 5 : 9])

plt.ylim(0, 1)
plt.xlim(0, 1)

ax5.plot([0.05, 1], [0.2, 0.2], color = '#5A8467')
ax5.scatter(0.05, 0.2, marker = 'o', color = '#5A8467')

# формируем границы диаграмм
ax5.spines['left'].set_visible(False)
ax5.spines['top'].set_visible(False)
ax5.spines['right'].set_visible(False)
ax5.spines['bottom'].set_visible(False)
ax5.set_facecolor('None')
ax5.patch.set_alpha(1)

# убираем оси графика
plt.xticks([])
plt.yticks([])

ax6 = plt.subplot(gs[3, 0 : 3])

plt.ylim(0, 1)
plt.xlim(0, 1)

ax6.plot([0, 1], [0.2, 0.2], color = '#606060')

# формируем границы диаграмм
ax6.spines['left'].set_visible(False)
ax6.spines['top'].set_visible(False)
ax6.spines['right'].set_visible(False)
ax6.spines['bottom'].set_visible(False)
ax6.set_facecolor('None')
ax6.patch.set_alpha(1)

# убираем оси графика
plt.xticks([])
plt.yticks([])

ax7 = plt.subplot(gs[3, 6 : 9])

plt.ylim(0, 1)
plt.xlim(0, 1)

ax7.plot([0, 1], [0.2, 0.2], color = '#606060')

# формируем границы диаграмм
ax7.spines['left'].set_visible(False)
ax7.spines['top'].set_visible(False)
ax7.spines['right'].set_visible(False)
ax7.spines['bottom'].set_visible(False)
ax7.set_facecolor('None')
ax7.patch.set_alpha(1)

# убираем оси графика
plt.xticks([])
plt.yticks([])

plt.savefig("system_photo/circle.png", dpi=300, bbox_inches='tight')
# plt.show()
#st.pyplot(fig)
st.image('system_photo/circle.png')

button = st.button('Сгенерировать презентацию')

if button:
    ppt = Presentation()  
    # задаем параметры слайдов (высота и ширина)
    ppt.slide_height = Inches(9) 
    ppt.slide_width = Inches(16)
    # за основу для слайдов берем пустой шаблон слайда
    blank_slide_layout = ppt.slide_layouts[6]  
    
    slide = ppt.slides.add_slide(blank_slide_layout)
    # задаем фон
    background = slide.shapes.add_picture(background_path, Cm(0), Cm(0), width=Inches(16), height=Inches(9))
    # рисуем иконку
    pic = slide.shapes.add_picture(icon_path, Cm(1.38), Cm(1.87), width=Cm(5.88), height=Cm(5.11))
    
    object_name = general['object_name'][0]
    
    month = general['date'][0].month
    if month // 10 == 0: 
        month = f'0{month}'

    day = general['date'][0].day
    if day // 10 == 0: 
        day = f'0{day}'
        
    # назначаем параметры текстовой ячейки заголовка слайда
    left = Inches(0.8503937007874016)
    top = Inches(0)
    width = Inches(16 - 2 * 0.8503937007874016)
    height = Inches(9)

    # создаем текстовую ячейку
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame

    # разрешаем перенос слов
    tf.word_wrap = True

    # добавляем параграф текста
    p = tf.add_paragraph()
    p.text = f'Доклад о ходе реализации проекта {general["object_type"][0][0 : len(general["object_type"][0]) - 1] + "я"} {object_name} для рассмотрения на Операционном комитете {day}.{month}.{general["date"][0].year}'

    # выравниваем по центрк по вертикали
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # выравниваем по центру по горизонтали
    p.alignment = PP_ALIGN.LEFT
    # делаем шрифт жирным, 40 размера, Arial Narrow, цвета (32, 56, 100)
    p.font.bold = True
    p.font.size = Pt(40)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    # назначаем параметры текстовой ячейки
    left = Cm(2.32)
    top = Cm(17.05)
    width = Cm(17.02)
    height = Cm(3.33)

    # создаем текстовую ячейку
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    # создаем первую строку "Докладчик"
    p = tf.add_paragraph()
    p.text = "Докладчик"
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(24)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(59, 56, 56)

    # создаем вторую строку с ФИО
    p1 = tf.add_paragraph()
    p1.text = general['fio'][0]
    p1.font.bold = True
    p1.alignment = PP_ALIGN.LEFT
    p1.font.size = Pt(24)
    p1.font.name = "Arial Narrow"
    p1.font.color.rgb = RGBColor(59, 56, 56)

    # создаем третью строку с должностью
    p2 = tf.add_paragraph()
    p2.text = general['job'][0]
    p2.alignment = PP_ALIGN.LEFT
    p2.font.size = Pt(24)
    p2.font.name = "Arial Narrow"
    p2.font.color.rgb = RGBColor(59, 56, 56)
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    ################################# График 1-ого уровня ########################################################################
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = "НА ВЕСЬ ПЕРИОД СТРОИТЕЛЬСТВА"
    name_of_slide(slide, f'ГРАФИК 1-ГО УРОВНЯ {object_name}', subinf)
    
    year = date.today().year
    
    pic = slide.shapes.add_picture("system_photo/1lvlgraph.png", Cm(6.56), Cm(3.05), width=Cm(33.06), height=Cm(18.54))
    pic = slide.shapes.add_picture("system_photo/legend1.png", Cm(1.02), Cm(21.2), width=Cm(25.4), height=Cm(0.75))
    
    left = Cm(1.02)
    top = Cm(9.5)
    width = Cm(6.55)
    height = Cm(4.1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = f'УТВЕРЖДЕННЫЙ ГРАФИК\n{general["object_type"][0][0 : len(general["object_type"][0]) - 1] + "я"} {object_name}'
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(18)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    left = Cm(1.02)
    top = Cm(15.57)
    width = Cm(4.63)
    height = Cm(1.8)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = f'ФАКТ/ПРОГНОЗ'
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(18)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    left = Cm(1.02)
    top = Cm(14.05)
    width = Cm(5.92)
    height = Cm(2.39)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = f'Отклонения от\nутвержденного графика'
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(16)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    #########################################################################################################################
    
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    # добавление названия слайда
    subinf = "С УКАЗАНИЕМ ФИЗИЧЕСКИХ ОБЪЕМОВ РАБОТ"
    name_of_slide(slide, f'КЛЮЧЕВЫЕ СОБЫТИЯ {general["date"][0].year} ГОДА', subinf)
    pic = slide.shapes.add_picture("system_photo/keyevents.png", Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
    pic = slide.shapes.add_picture("system_photo/legend2.png", Cm(1.25), Cm(22.11), width=Cm(15.06), height=Cm(0.73))
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = f'ПО ПРОЕКТУ {object_name}'
    name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПЛАНА ПО ОСВОЕНИЮ {general["date"][0].year} ГОДА', subinf)
    
    left = Cm(0.29)
    top = Cm(1.76)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (ОСВОЕНИЕ), млн. долл."
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    top = Cm(4.05)
    table = draw_table(shapes, top)

    table.cell(3, 2).text = "100%"
    table.cell(4, 2).text = "100%"
    table.cell(5, 2).text = "100%"
    table.cell(6, 2).text = "100%"
    table.cell(7, 2).text = "100%"
    table.cell(8, 2).text = "100%"
    table.cell(9, 2).text = "100%"

    table.cell(3, 3).text = "662.3"
    table.cell(4, 3).text = "530.5"
    table.cell(5, 3).text = "114.7"
    table.cell(6, 3).text = "7.4"
    table.cell(7, 3).text = "9.7"
    table.cell(8, 3).text = "151.7"
    table.cell(9, 3).text = str(float(table.cell(3, 3).text) + float(table.cell(8, 3).text))

    table.cell(3, 4).text = "42.1"
    table.cell(4, 4).text = "27.8"
    table.cell(5, 4).text = "14.3"
    table.cell(6, 4).text = "0"
    table.cell(7, 4).text = "0"
    table.cell(8, 4).text = "23.8"
    table.cell(9, 4).text = str(float(table.cell(3, 4).text) + float(table.cell(8, 4).text))

    table.cell(3, 5).text = "56.1"
    table.cell(4, 5).text = "33.5"
    table.cell(5, 5).text = "22.6"
    table.cell(6, 5).text = "0"
    table.cell(7, 5).text = "0"
    table.cell(8, 5).text = "21.9"
    table.cell(9, 5).text = str(float(table.cell(3, 5).text) + float(table.cell(8, 5).text))

    table.cell(3, 6).text = "(+42.4)"
    table.cell(4, 6).text = "(+28.6)"
    table.cell(5, 6).text = "(+13.8)"
    table.cell(6, 6).text = "0"
    table.cell(7, 6).text = "0"
    table.cell(8, 6).text = "(+6.5)"
    table.cell(9, 6).text = "(+48.9)"

    table.cell(3, 7).text = "133%"
    table.cell(4, 7).text = "121%"
    table.cell(5, 7).text = "158%"
    table.cell(6, 7).text = "-"
    table.cell(7, 7).text = "-"
    table.cell(8, 7).text = "92%"
    table.cell(9, 7).text = "118%"

    table.cell(3, 8).text = "(149%)"
    table.cell(4, 8).text = "(125%)"
    table.cell(5, 8).text = "(>100%)"
    table.cell(6, 8).text = "-"
    table.cell(7, 8).text = "-"
    table.cell(8, 8).text = "(77%)"
    table.cell(9, 8).text = "(133%)"

    table.cell(3, 9).text = "8%"
    table.cell(4, 9).text = "6%"
    table.cell(5, 9).text = "20%"
    table.cell(6, 9).text = "0%"
    table.cell(7, 9).text = "0%"
    table.cell(8, 9).text = "14%"
    table.cell(9, 9).text = "10%"

    table.cell(3, 10).text = "100%"
    table.cell(4, 10).text = "100%"
    table.cell(5, 10).text = "100%"
    table.cell(6, 10).text = "100%"
    table.cell(7, 10).text = "100%"
    table.cell(8, 10).text = "100%"
    table.cell(9, 10).text = "100%"

    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            if cnt <= 2:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1
    
    pic = slide.shapes.add_picture("system_photo/plotplancompl.png", Inches(0.161417), Inches(4.77), width=Inches(11.46), height=Inches(3.85))
    pic = slide.shapes.add_picture("system_photo/bar1.png", Inches(11.96), Inches(4.62), width=Inches(3.84), height=Inches(2.19))
    
    # For adjusting the  Margins in inches 
    left = Cm(0.77)
    top = Cm(21.26)
    height = Cm(1.45)
    width = Cm(38.85)
              
    # creating textBox 
    txBox = slide.shapes.add_textbox(left, top, 
                                             width, height) 
              
    # creating textFrames 
    tf = txBox.text_frame 
    # adding Paragraphs
    p = tf.add_paragraph()  
              
    # adding text 
    p.text = "Какие-то сноски"
              
    # font  
    p.font.name = "Arial Narrow"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(89, 89, 89)
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = f'ПО ПРОЕКТУ {object_name}'
    name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПЛАНА ПО ОСВОЕНИЮ {general["date"][0].year} ГОДА ПО СТРУКТУРЕ ЗАТРАТ', subinf)
    
    pic = slide.shapes.add_picture("system_photo/4plot.png", Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
    
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = f''
    name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПОКАЗАТЕЛЯ «ОСВОЕНИЕ» ПО ПРОЕКТУ {object_name} В РАЗРЕЗЕ ЭНЕРГОБЛОКОВ АЭС', subinf)
    
    left = Cm(0.29)
    top = Cm(1.76)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "Блок №1"
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    top = Cm(4.05)
    table = draw_table(shapes, top)

    table.cell(3, 2).text = "100%"
    table.cell(4, 2).text = "100%"
    table.cell(5, 2).text = "100%"
    table.cell(6, 2).text = "100%"
    table.cell(7, 2).text = "100%"
    table.cell(8, 2).text = "100%"
    table.cell(9, 2).text = "100%"

    table.cell(3, 3).text = "662.3"
    table.cell(4, 3).text = "530.5"
    table.cell(5, 3).text = "114.7"
    table.cell(6, 3).text = "7.4"
    table.cell(7, 3).text = "9.7"
    table.cell(8, 3).text = "151.7"
    table.cell(9, 3).text = str(float(table.cell(3, 3).text) + float(table.cell(8, 3).text))

    table.cell(3, 4).text = "42.1"
    table.cell(4, 4).text = "27.8"
    table.cell(5, 4).text = "14.3"
    table.cell(6, 4).text = "0"
    table.cell(7, 4).text = "0"
    table.cell(8, 4).text = "23.8"
    table.cell(9, 4).text = str(float(table.cell(3, 4).text) + float(table.cell(8, 4).text))

    table.cell(3, 5).text = "56.1"
    table.cell(4, 5).text = "33.5"
    table.cell(5, 5).text = "22.6"
    table.cell(6, 5).text = "0"
    table.cell(7, 5).text = "0"
    table.cell(8, 5).text = "21.9"
    table.cell(9, 5).text = str(float(table.cell(3, 5).text) + float(table.cell(8, 5).text))

    table.cell(3, 6).text = "(+42.4)"
    table.cell(4, 6).text = "(+28.6)"
    table.cell(5, 6).text = "(+13.8)"
    table.cell(6, 6).text = "0"
    table.cell(7, 6).text = "0"
    table.cell(8, 6).text = "(+6.5)"
    table.cell(9, 6).text = "(+48.9)"

    table.cell(3, 7).text = "133%"
    table.cell(4, 7).text = "121%"
    table.cell(5, 7).text = "158%"
    table.cell(6, 7).text = "-"
    table.cell(7, 7).text = "-"
    table.cell(8, 7).text = "92%"
    table.cell(9, 7).text = "118%"

    table.cell(3, 8).text = "(149%)"
    table.cell(4, 8).text = "(125%)"
    table.cell(5, 8).text = "(>100%)"
    table.cell(6, 8).text = "-"
    table.cell(7, 8).text = "-"
    table.cell(8, 8).text = "(77%)"
    table.cell(9, 8).text = "(133%)"

    table.cell(3, 9).text = "8%"
    table.cell(4, 9).text = "6%"
    table.cell(5, 9).text = "20%"
    table.cell(6, 9).text = "0%"
    table.cell(7, 9).text = "0%"
    table.cell(8, 9).text = "14%"
    table.cell(9, 9).text = "10%"

    table.cell(3, 10).text = "100%"
    table.cell(4, 10).text = "100%"
    table.cell(5, 10).text = "100%"
    table.cell(6, 10).text = "100%"
    table.cell(7, 10).text = "100%"
    table.cell(8, 10).text = "100%"
    table.cell(9, 10).text = "100%"

    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            if cnt <= 2:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1
    
    left = Cm(0.29)
    top = Cm(10.74)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "Блок №2"
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    top = Cm(12.95)
    table = draw_table(shapes, top)

    table.cell(3, 2).text = "100%"
    table.cell(4, 2).text = "100%"
    table.cell(5, 2).text = "100%"
    table.cell(6, 2).text = "100%"
    table.cell(7, 2).text = "100%"
    table.cell(8, 2).text = "100%"
    table.cell(9, 2).text = "100%"

    table.cell(3, 3).text = "662.3"
    table.cell(4, 3).text = "530.5"
    table.cell(5, 3).text = "114.7"
    table.cell(6, 3).text = "7.4"
    table.cell(7, 3).text = "9.7"
    table.cell(8, 3).text = "151.7"
    table.cell(9, 3).text = str(float(table.cell(3, 3).text) + float(table.cell(8, 3).text))

    table.cell(3, 4).text = "42.1"
    table.cell(4, 4).text = "27.8"
    table.cell(5, 4).text = "14.3"
    table.cell(6, 4).text = "0"
    table.cell(7, 4).text = "0"
    table.cell(8, 4).text = "23.8"
    table.cell(9, 4).text = str(float(table.cell(3, 4).text) + float(table.cell(8, 4).text))

    table.cell(3, 5).text = "56.1"
    table.cell(4, 5).text = "33.5"
    table.cell(5, 5).text = "22.6"
    table.cell(6, 5).text = "0"
    table.cell(7, 5).text = "0"
    table.cell(8, 5).text = "21.9"
    table.cell(9, 5).text = str(float(table.cell(3, 5).text) + float(table.cell(8, 5).text))

    table.cell(3, 6).text = "(+42.4)"
    table.cell(4, 6).text = "(+28.6)"
    table.cell(5, 6).text = "(+13.8)"
    table.cell(6, 6).text = "0"
    table.cell(7, 6).text = "0"
    table.cell(8, 6).text = "(+6.5)"
    table.cell(9, 6).text = "(+48.9)"

    table.cell(3, 7).text = "133%"
    table.cell(4, 7).text = "121%"
    table.cell(5, 7).text = "158%"
    table.cell(6, 7).text = "-"
    table.cell(7, 7).text = "-"
    table.cell(8, 7).text = "92%"
    table.cell(9, 7).text = "118%"

    table.cell(3, 8).text = "(149%)"
    table.cell(4, 8).text = "(125%)"
    table.cell(5, 8).text = "(>100%)"
    table.cell(6, 8).text = "-"
    table.cell(7, 8).text = "-"
    table.cell(8, 8).text = "(77%)"
    table.cell(9, 8).text = "(133%)"

    table.cell(3, 9).text = "8%"
    table.cell(4, 9).text = "6%"
    table.cell(5, 9).text = "20%"
    table.cell(6, 9).text = "0%"
    table.cell(7, 9).text = "0%"
    table.cell(8, 9).text = "14%"
    table.cell(9, 9).text = "10%"

    table.cell(3, 10).text = "100%"
    table.cell(4, 10).text = "100%"
    table.cell(5, 10).text = "100%"
    table.cell(6, 10).text = "100%"
    table.cell(7, 10).text = "100%"
    table.cell(8, 10).text = "100%"
    table.cell(9, 10).text = "100%"

    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            if cnt <= 2:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1
    
    # For adjusting the  Margins in inches 
    left = Cm(0.77)
    top = Cm(21.26)
    height = Cm(1.45)
    width = Cm(38.85)
              
    # creating textBox 
    txBox = slide.shapes.add_textbox(left, top, 
                                             width, height) 
              
    # creating textFrames 
    tf = txBox.text_frame 
    # adding Paragraphs
    p = tf.add_paragraph()  
              
    # adding text 
    p.text = "Какие-то сноски"
              
    # font  
    p.font.name = "Arial Narrow"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(89, 89, 89)
    
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = f'ПО ПРОЕКТУ {object_name}'
    name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПЛАНА ПО РЕАЛИЗАЦИИ {general["date"][0].year} ГОДА', subinf)
    
    left = Cm(0.29)
    top = Cm(1.76)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (РЕАЛИЗАЦИЯ), млн. долл."
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    top = Cm(4.05)
    table = draw_table(shapes, top)

    table.cell(3, 2).text = "100%"
    table.cell(4, 2).text = "100%"
    table.cell(5, 2).text = "100%"
    table.cell(6, 2).text = "100%"
    table.cell(7, 2).text = "100%"
    table.cell(8, 2).text = "100%"
    table.cell(9, 2).text = "100%"

    table.cell(3, 3).text = "662.3"
    table.cell(4, 3).text = "530.5"
    table.cell(5, 3).text = "114.7"
    table.cell(6, 3).text = "7.4"
    table.cell(7, 3).text = "9.7"
    table.cell(8, 3).text = "151.7"
    table.cell(9, 3).text = str(float(table.cell(3, 3).text) + float(table.cell(8, 3).text))

    table.cell(3, 4).text = "42.1"
    table.cell(4, 4).text = "27.8"
    table.cell(5, 4).text = "14.3"
    table.cell(6, 4).text = "0"
    table.cell(7, 4).text = "0"
    table.cell(8, 4).text = "23.8"
    table.cell(9, 4).text = str(float(table.cell(3, 4).text) + float(table.cell(8, 4).text))

    table.cell(3, 5).text = "56.1"
    table.cell(4, 5).text = "33.5"
    table.cell(5, 5).text = "22.6"
    table.cell(6, 5).text = "0"
    table.cell(7, 5).text = "0"
    table.cell(8, 5).text = "21.9"
    table.cell(9, 5).text = str(float(table.cell(3, 5).text) + float(table.cell(8, 5).text))

    table.cell(3, 6).text = "(+42.4)"
    table.cell(4, 6).text = "(+28.6)"
    table.cell(5, 6).text = "(+13.8)"
    table.cell(6, 6).text = "0"
    table.cell(7, 6).text = "0"
    table.cell(8, 6).text = "(+6.5)"
    table.cell(9, 6).text = "(+48.9)"

    table.cell(3, 7).text = "133%"
    table.cell(4, 7).text = "121%"
    table.cell(5, 7).text = "158%"
    table.cell(6, 7).text = "-"
    table.cell(7, 7).text = "-"
    table.cell(8, 7).text = "92%"
    table.cell(9, 7).text = "118%"

    table.cell(3, 8).text = "(149%)"
    table.cell(4, 8).text = "(125%)"
    table.cell(5, 8).text = "(>100%)"
    table.cell(6, 8).text = "-"
    table.cell(7, 8).text = "-"
    table.cell(8, 8).text = "(77%)"
    table.cell(9, 8).text = "(133%)"

    table.cell(3, 9).text = "8%"
    table.cell(4, 9).text = "6%"
    table.cell(5, 9).text = "20%"
    table.cell(6, 9).text = "0%"
    table.cell(7, 9).text = "0%"
    table.cell(8, 9).text = "14%"
    table.cell(9, 9).text = "10%"

    table.cell(3, 10).text = "100%"
    table.cell(4, 10).text = "100%"
    table.cell(5, 10).text = "100%"
    table.cell(6, 10).text = "100%"
    table.cell(7, 10).text = "100%"
    table.cell(8, 10).text = "100%"
    table.cell(9, 10).text = "100%"

    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            if cnt <= 2:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1
       
    pic = slide.shapes.add_picture("system_photo/plot2.png", Cm(2.71), top = Cm(11.53), width=Cm(34.96), height=Cm(11.18))
    
    # For adjusting the  Margins in inches 
    left = Cm(0.77)
    top = Cm(21.26)
    height = Cm(1.45)
    width = Cm(38.85)
              
    # creating textBox 
    txBox = slide.shapes.add_textbox(left, top, 
                                             width, height) 
              
    # creating textFrames 
    tf = txBox.text_frame 
    # adding Paragraphs
    p = tf.add_paragraph()  
              
    # adding text 
    p.text = "Какие-то сноски"
              
    # font  
    p.font.name = "Arial Narrow"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(89, 89, 89)
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    
    background13 = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(16), Inches(9))
    fill = background13.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    background13.line.color.rgb = RGBColor(255, 255, 255)
    shadow = background13.shadow
    shadow.inherit = False
    
    rec = shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(10.44), Cm(0), Cm(2.53), Cm(11.18))
    fill = rec.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(58, 123, 168)
    rec.line.color.rgb = RGBColor(58, 123, 168)
    shadow = rec.shadow
    shadow.inherit = False
    
    left = Cm(13.8)
    top = Cm(7.78)
    width = Cm(12.18)
    height = Cm(3.59)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "ПРИЛОЖЕНИЯ"
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.bold = True
    p.font.size = Pt(54)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(58, 123, 168)
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = ""
    name_of_slide(slide, f'СТАТУС ВЫДАЧИ РД НА ОБЪЕМ СМР {general["date"][0].year} ГОДУ', subinf)
    
    pic = slide.shapes.add_picture("system_photo/2bigplot.png", Inches(0.401575), Cm(3.12), width=Inches(15.2), height=Inches(7.4))
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = ""
    name_of_slide(slide, f'ЧИСЛЕННОСТЬ СТРОИТЕЛЬНОГО ПЕРСОНАЛА НА ПЛОЩАДКЕ В {general["date"][0].year} ГОДУ', subinf)
    
    pic = slide.shapes.add_picture("system_photo/bigtable.png", Inches(0.401575), Inches(1.4), width=Inches(15.2), height=Inches(7.4))
    
    left = Cm(0.29)
    top = Cm(1.76)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "ЧИСЛЕННОСТЬ СТРОИТЕЛЬНОГО ПЕРСОНАЛА, чел."
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = ""
    name_of_slide(slide, f'ИНВЕНТАРИЗАЦИЯ ПРОЕКТА СООРУЖЕНИЕ {object_name} НА ПРЕДМЕТ НАЛИЧИЯ РИСКОВ СРЫВА ПОСТАВОК ОБОРУДОВАНИЯ И КОМПЛЕКТУЮЩИХ ИЗ 3-Х СТРАН', subinf)
    
    pic = slide.shapes.add_picture("system_photo/circle.png", Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
    
    # создаем слайд
    slide = ppt.slides.add_slide(blank_slide_layout)
    shapes = slide.shapes
    subinf = ""
    name_of_slide(slide, f'СТАТУС ВЫПОЛНЕНИЯ ПОРУЧЕНИЙ ГЕНЕРАЛЬНОГО ДИРЕКТОРА В РАМКАХ ПРОЕКТА СТРОИТЕЛЬСТВА', subinf)
    
    left = Cm(0.29)
    top = Cm(1.76)
    width = Cm(39.84)
    height = Cm(2.48)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Enable word wrap
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = "СТАТУС ВЫПОЛНЕНИЯ ПОРУЧЕНИЙ ГЕНЕРАЛЬНОГО ДИРЕКТОРА В РАМКАХ ПРОЕКТА СТРОИТЕЛЬСТВА"
    # Optionally, if you want to center the entire textbox including its vertical position
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Center the paragraph text
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(20)
    p.font.name = "Arial Narrow"
    p.font.color.rgb = RGBColor(32, 56, 100)
    
    rows = 3
    cols = 6
    left = Cm(0.41)
    top = Cm(4.05)
    width = Cm(39.63)
    height = Cm(7.54)

    table = shapes.add_table(rows, cols, left, top, width, height).table

    table.cell(0, 0).text = "№"
    table.cell(0, 1).text = "ПОРУЧЕНИЕ"
    table.cell(0, 2).text = "СРОК ВЫПОЛНЕНИЯ"
    table.cell(0, 3).text = "ОТВ."
    table.cell(0, 4).text = "СТАТУС (выполнено / не выполнено / срок не наступил)"
    table.cell(0, 5).text = "КОММЕНТАРИИ"

    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "7.1. Дерию А.В., Зотеевой А.Г., Шперле О.Н. совместно с Петровым А.Ю., Комаровым К.Б. и Сахаровым Г.С. обеспечить максимальную концентрацию усилий по достижению выполнения событий, связанных с физическим пуском соответствующих блоков по проектам сооружения АЭС «Аккую», АЭС «Руппур» и Курской АЭС в соответствии с утвержденными сроками по их реализации. Дерию А.В., Зотеевой А.Г., Шперле О.Н. в рамках докладов о ходе реализации проектов к заседаниям Операционного комитета Госкорпорации «Росатом» уделять особое внимание прогнозу исполнения ключевого события «Физический пуск» и рисках его выполнения. (Протокол заседания Операционного комитета ГК «Росатом» от 15.11.2023 №1-ОК/114-Пр)."
    table.cell(1, 2).text = "в рамках докладов о ходе реализации проектов к заседаниям Операционного комитета Госкорпорации «Росатом»"
    table.cell(1, 3).text = "Дерий А.В., Зотеева А.Г., Шперле О.Н., Петров А.Ю."
    table.cell(1, 4).text = "Будет учтено в рамках доклада на текущем и последующих заседаниях Операционного комитета Госкорпорации «Росатом»"
    table.cell(1, 5).text = "Контроль за выполнением поставленной задачи по обеспечению физического пуска  Блока 1 АЭС «Руппур» в 2024 году, а также принятие всех необходимый решений проводится в рамках ежемесячных штабов под председательством А.Ю. Петрова."

    table.cell(2, 0).text = "2"
    table.cell(2, 1).text = "2.2. Дерию А.В. по согласованию с Петровым А.Ю. и Волковым Д.А. к следующему очередному тематическому заседанию Операционного комитета ГК «Росатом» откорректировать доклад о ходе реализации проекта сооружения АЭС «Руппур», блоки 1-2, в части детализации этапов и событий, находящихся на критическом пути реализации проекта в 2024 - 2025 г.г. (Протокол заседания Операционного комитета ГК «Росатом» от 18.04.2024 №1-ОК/23-Пр)."
    table.cell(2, 2).text = "13.05.2024"
    table.cell(2, 3).text = "Дерий А.В."
    table.cell(2, 4).text = "Выполнено"

    cnt = 0
    # set column widths
    for col in table.columns:
        if cnt == 1:
            col.width = Cm(11.5)
        elif cnt == 0:
            col.width = Cm(1.5)
        cnt += 1
    cnt = 0
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(12)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
            
            fill = cell.fill
            fill.solid()
            if cnt == 0:
                fill.fore_color.rgb = RGBColor(218, 232, 242)
            else:
                fill.fore_color.rgb = RGBColor(255, 255, 255)
        cnt += 1
    
    
    cnt = 0
    for slide in ppt.slides:
        if cnt != 0 and cnt != 7:
            line = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(8.7), Inches(15.2), Inches(0))
            line.line.color.rgb = RGBColor(230, 230, 230)
            shadow = line.shadow
            shadow.inherit = False

            # For adjusting the  Margins in inches 
            left = Inches(15.62992125984252) - Pt(10) * ((cnt + 1) // 10)
            top = Inches(8.295275590551181)
            height = width = Inches(1)
              
            # creating textBox 
            txBox = slide.shapes.add_textbox(left, top, 
                                             width, height) 
              
            # creating textFrames 
            tf = txBox.text_frame 
            # adding Paragraphs 
            p = tf.add_paragraph()  
              
            # adding text 
            p.text = f'{cnt + 1}' 
              
            # font  
            p.font.bold = True
            p.font.name = "Arial Narrow"
            p.font.size = Pt(20)
            p.font.color.rgb = RGBColor(89, 89, 89)


            line2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.16), Inches(15.2), Inches(0))
            line2.line.color.rgb = RGBColor(26, 60, 123)
            shadow = line2.shadow
            shadow.inherit = False

            pic = slide.shapes.add_picture(icon2_path, Cm(0.41), Cm(0.35), width=Cm(2.38), height=Cm(2.53))
        cnt += 1
        
    # сохранение
    ppt.save(f"ppt.pptx")
    st.write('Генерация прошла успешно!')


# ## Статус выполнения поручений ген директора

# ## Оформление для всех страниц (нумерация и т. д.)

# # Сохранение и закрытие презентации
