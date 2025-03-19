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
def draw_foreign_revenue_graph_accum(general, foreign_revenue_accum):
    f, ax = plt.subplots(1, 1)
    f.set_size_inches(11.46, 3.85)
    plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
    x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    plan_y = np.array(foreign_revenue_accum.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(foreign_revenue_accum.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([foreign_revenue_accum.iloc[2, 11]])
    month = general['date'][0].month
    make_graph(ax, "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (НАКОПИТЕЛЬНО), млн. долл.", x, plan_y, fact_y, pred_y, month)
    plt.savefig("system_photo/foreign_revenue_graph.png", dpi=300, bbox_inches='tight')