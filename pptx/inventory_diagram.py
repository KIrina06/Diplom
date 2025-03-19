### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт из файлов
from func import make_width, make_height, mini_func, draw_text_center, draw_text_top, draw_text

### функции
@st.cache_resource
def draw_inventory_diagram(inventory):
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
                                      autopct=lambda pct: mini_func(pct),
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
    fontsize = 44
    
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
    
    plt.savefig("system_photo/inventory_diagram.png", dpi=300, bbox_inches='tight')