### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт файлов
from func import min_date, max_date, count_coords

# функция отрисовки графика первого уровня
@st.cache_resource
def draw_graph_1_lvl(general, graph_1_lvl):
    # получаем список всех годов
    years = list()
    
    for i in graph_1_lvl['date_plan']:
        if years.count(i.year) == 0:
            years.append(i.year)
    
    # получаем веса всех годов
    years_weights = list([0] * len(years))
    
    for i in graph_1_lvl['date_plan']:
        for j in range(len(years)):
            if i.year == years[j]:
                years_weights[j] += 1
    
    # делаем словарь с годами и весами, как года будут отображаться на графике
    years_dict = {}
    temp_dict = {}
    year_1 = general['date'][0].year - 1
    year_0 = year_1
    year_0_1_weight = 0
    
    for i in range(len(years) - 1, -1, -1):
        if years[i] >= general['date'][0].year:
            temp_dict = {years[i]: years_weights[i]}
            years_dict, temp_dict = temp_dict, years_dict
            years_dict.update(temp_dict)
        else:
            if years[i] < year_0:
                year_0 = years[i]
            year_0_1_weight += years_weights[i]
    temp_dict = {f'{year_0} - {year_1}': year_0_1_weight}
    years_dict, temp_dict = temp_dict, years_dict
    years_dict.update(temp_dict)
    
    # определяем, сколько места будет занимать каждый год
    
    # ищем сумму всех весов
    sum_weights = 0
    
    for k, v in years_dict.items():
        sum_weights += v
    
    # координаты, разделяющие один год от другого
    years_coords = list()
    years_coords.append(0)
    
    for k, v in years_dict.items():
        years_coords.append(v / sum_weights * 100 + years_coords[len(years_coords) - 1])
        
    # распределяем события по годам
    events_dict = {}
    
    for k, v in years_dict.items():
        events_dict[k] = []
    
    for i in range(len(graph_1_lvl['date_plan'])):
        for k, v in events_dict.items():
            if graph_1_lvl['date_plan'][i].year == k:
                events_dict[k].extend([{'event_name': graph_1_lvl['event_name'][i], 'date_plan': graph_1_lvl['date_plan'][i], 
                                      'date_fact_forecast': graph_1_lvl['date_fact_forecast'][i], 'x_plan': [], 'x_fact_forecast': []}])
            elif graph_1_lvl['date_plan'][i].year < general['date'][0].year:
                events_dict[list(events_dict.keys())[0]].extend([{'event_name': graph_1_lvl['event_name'][i], 'date_plan': graph_1_lvl['date_plan'][i], 
                                      'date_fact_forecast': graph_1_lvl['date_fact_forecast'][i], 'x_plan': [], 'x_fact_forecast': []}])
                break;
    
    # расчет координат x_plan
    tmp = 0
    for k, v in events_dict.items():
        events = events_dict[k]
        x = count_coords(events, years_dict, sum_weights, k, 'date_plan')
        for i in range(len(x)):
            x[i] += years_coords[tmp]
            events_dict[k][i]['x_plan'] = x[i]
        tmp += 1
    
    tmp = 0
    
    # расчет координат x_fact_forecast
    for k, v in events_dict.items():
        events = events_dict[k]
        x = count_coords(events, years_dict, sum_weights, k, 'date_fact_forecast')
        for i in range(len(x)):
            x[i] += years_coords[tmp]
            events_dict[k][i]['x_fact_forecast'] = x[i]
        tmp += 1

    # строим сам график
    plt.figure(figsize=(20,10))
    
    # разметка
    plt.xlim(0, 102)
    plt.ylim(0, 4)
    
    # скрытие осей
    ax = plt.gca()
    plt.axis('off')
    
    # рисуем сами временные линии
    plt.plot([0, 100.2], [1, 1], [0, 100.2], [2, 2], [0, 100.2], [3.8, 3.8], color="#BDD7EE", linewidth=30, zorder = 1)
    plt.arrow(x=0, y=1, dx=101.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)
    plt.arrow(x=0, y=2, dx=101.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)
    plt.arrow(x=0, y=3.8, dx=101.5, dy=0, width=0.08, facecolor='#BDD7EE', edgecolor='none', zorder = 0)
    
    # нанесем года и разграничения
    # разграничения
    for i in range(1, len(years_coords) - 1):
        plt.plot(years_coords[i], 3.8, marker='|', markersize=30, markerfacecolor='w', markeredgecolor='w', zorder = 2)
    
    # метки годов
    for i in range(len(years_coords) - 1):
        plt.text((years_coords[i] + years_coords[i + 1])/2, 3.8, f'{list(years_dict.keys())[i]}', ha='center', va='center', color='#003274', fontsize = 16, fontweight = 'bold')
    
    # наносим все точки и подписи со стрелочками
    for k, v in events_dict.items():
        for i in range(len(v)):
            plt.scatter(v[i]['x_plan'], 2, marker='o', facecolor='none', edgecolor='#2E75B6', s=800, linewidths=7, zorder = 2)
            
            if v[i]['date_fact_forecast'] > general['date'][0]:
                plt.scatter(v[i]['x_fact_forecast'], 1, marker='o', facecolor='none', linestyle=':', edgecolor='#1F435F', s=800, linewidths=7, zorder = 2)
            else:
                if v[i]['date_fact_forecast'] > v[i]['date_plan']:
                    plt.scatter(v[i]['x_fact_forecast'], 1, marker='o', facecolor='none', edgecolor='red', s=800, linewidths=7, zorder = 2)
                else:
                    plt.scatter(v[i]['x_fact_forecast'], 1, marker='o', facecolor='none', edgecolor='#00B050', s=800, linewidths=7, zorder = 2)
            
            plt.text(v[i]['x_plan'], 1.6, f'{v[i]['date_plan'].day}.{v[i]['date_plan'].month}.{v[i]['date_plan'].year}', va = 'center', ha = 'center', rotation = 45, color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
            plt.text(v[i]['x_fact_forecast'], 0.6, f'{v[i]['date_fact_forecast'].day}.{v[i]['date_fact_forecast'].month}.{v[i]['date_fact_forecast'].year}', va = 'center', ha = 'center', rotation = 45, color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
    
            if len(v[i]['event_name']) > 32:
                v[i]['event_name'] = v[i]['event_name'][0 : 30] + "..."
            plt.text(v[i]['x_plan'], 2.9, v[i]['event_name'], va = 'center', ha = 'center', rotation = 45, color = '#003274', fontsize = 12, bbox = dict(boxstyle = 'square', fc = 'w', ec = '#003274', lw = 1))
            plt.annotate("", xy=(v[i]['x_plan'], 2), xytext=(v[i]['x_plan'], 2.9), arrowprops=dict(arrowstyle="->", color = '#003274', linestyle = 'dashed', linewidth = 0.7), zorder = 1)
            plt.annotate("", xy=(v[i]['x_fact_forecast'], 1), xytext=(v[i]['x_plan'], 2), arrowprops=dict(arrowstyle="->", color = '#003274', linewidth = 2), zorder = 4)
    
    # добавление флажка
    key = list(events_dict.keys())[len(list(events_dict.keys())) - 1]
    x_flag = events_dict[key][len(events_dict[key]) - 1]['x_plan']
    y_flag = 2.2
    plt.plot(x_flag, y_flag, marker='|', markersize=60, markerfacecolor='#2E75B6', markeredgecolor='#2E75B6', zorder = 3)
    plt.plot(x_flag + 1, y_flag + 0.19, marker='>', markersize=25, markerfacecolor='#2E75B6', markeredgecolor='#2E75B6', zorder = 3)
    plt.plot(x_flag + 1, y_flag + 0.12, marker='>', markersize=25, markerfacecolor='#2E75B6', markeredgecolor='none', zorder = 3)

    # добавление текущей даты
    tmp = 0
    for k, v in years_dict.items():
        tmp += 1
        if k == general['date'][0].year:
            break
    x_left = years_coords[tmp - 1]  
    x_right = years_coords[tmp]  

    plt.plot([x_left, x_right], [0.3, 0.3], linestyle = '--', linewidth = 3, color = '#00B050', zorder = 1)
    plt.plot([x_left, x_right], [3.91, 3.91], linestyle = '--', linewidth = 3, color = '#00B050', zorder = 1)
    plt.plot([x_left, x_left], [0.3, 3.91], linestyle = '--', linewidth = 3, color = '#00B050', zorder = 1)
    plt.plot([x_right, x_right], [0.3, 3.91], linestyle = '--', linewidth = 3, color = '#00B050', zorder = 1)

    dates = [pd.Timestamp(f'{general['date'][0].year}-{1}-{1} 00:00:00'), general['date'][0], pd.Timestamp(f'{general['date'][0].year}-{12}-{31} 00:00:00')]
    date_min = dates[0].year % 100 + dates[0].month / 12 * 2 + dates[0].day / 300 * 2
    date_now = dates[1].year % 100 + dates[1].month / 12 * 2 + dates[1].day / 300 * 2
    date_max = dates[2].year % 100 + dates[2].month / 12 * 2 + dates[2].day / 300 * 2
    date_now = (date_now - date_min) / (date_max - date_min) * (years_dict[general['date'][0].year] / sum_weights * 100) + x_left

    plt.plot([date_now, date_now], [0.2, 4], linestyle = '--', linewidth = 3, color = '#00B050', zorder = 1)
    plt.text(date_now, 0.09, 'текущая дата', fontsize = 14, fontfamily = 'cursive', ha = 'center', va = 'bottom', color = '#00B050')
    
    plt.savefig("system_photo/1lvlgraph.png", dpi=300, bbox_inches='tight')