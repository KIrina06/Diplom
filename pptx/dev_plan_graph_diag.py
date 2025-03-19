### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт файлов
from func import make_graph

# диаграмма графика выполнения программы накопительно
@st.cache_resource
def draw_dev_plan_graph_accum(general, accumulative_execution):
    f, ax = plt.subplots(1, 1)
    f.set_size_inches(11.46, 3.85)
    plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
    x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    plan_y = np.array(accumulative_execution.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(accumulative_execution.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([accumulative_execution.iloc[2, 11]])
    month = general['date'][0].month
    make_graph(ax, "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (НАКОПИТЕЛЬНО), млн. долл.", x, plan_y, fact_y, pred_y, month)
    plt.savefig("system_photo/dev_plan_graph_accum.png", dpi=300, bbox_inches='tight')

# диаграмма численности строит. персонала
def draw_dev_plan_diag_1(dev_plan_diag_1):
    plan_dso = dev_plan_diag_1['dso']
    plan_s = dev_plan_diag_1['other']
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
    plt.savefig("system_photo/dev_plan_diag_1.png", dpi=300, bbox_inches='tight')

# диаграмма выдачи рд
def draw_dev_plan_diag_2(dev_plan_diag_2):
    proc = dev_plan_diag_2['proc']
    compl = dev_plan_diag_2['compl']
    x = ['ПЛАН', 'ФАКТ']
    
    fig = plt.figure(figsize=(3, 1))
    ax = fig.add_subplot()
    plt.ylim(0, proc[0] + proc[1])
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    ax.set_yticklabels([])
    ax.tick_params(axis='y', length=0)
    
    ax.bar(x, proc, color="#bfbfbf")
    
    for i, v in enumerate(proc):
        ax.text(i, proc[i], f"{proc[i]}%\n({compl[i]} компл.)", ha="center", va="bottom", fontsize=10, color="#595959", fontweight="bold")
    
    plt.title(f"ВЫДАНО РД НА ОБЪЕМ СМР\nТЕКУЩЕГО ГОДА", fontsize=12, color="#1a3c7b", fontweight="semibold")
    plt.savefig("system_photo/dev_plan_diag_2.png", dpi=300, bbox_inches='tight')