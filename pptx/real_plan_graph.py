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
def draw_real_plan_graph_accum(general, accumulative_realization):
    f, ax = plt.subplots(1, 1)
    f.set_size_inches(11.46, 3.85)
    plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
    x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    plan_y = np.array(accumulative_realization.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(accumulative_realization.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([accumulative_realization.iloc[2, 11]])
    month = general['date'][0].month
    make_graph(ax, "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (НАКОПИТЕЛЬНО), млн. долл.", x, plan_y, fact_y, pred_y, month)
    plt.savefig("system_photo/real_plan_graph.png", dpi=300, bbox_inches='tight')