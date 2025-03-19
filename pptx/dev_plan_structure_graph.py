### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd
import streamlit as st

### импорт из файлов
from func import make_graph

### функции

# графики по выполнению плана по структуре затрат
@st.cache_resource
def draw_dev_plan_structure(general, pir, equipment, smr, other):
    f, ax = plt.subplots(2, 2)
    f.set_size_inches(15.2, 7.4)
    plt.subplots_adjust(wspace = 0.25, hspace = 0.25)
    x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    plan_y = np.array(pir.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(pir.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([pir.iloc[2, 11]])
    month = general['date'][0].month
    make_graph(ax[0, 0], "ПИР", x, plan_y, fact_y, pred_y, month)
    
    x = ['ЯНВ', 'ФЕВ', 'МАР', 'АПР', 'МАЙ', 'ИЮН', 'ИЮЛ', 'АВГ', 'СЕН', 'ОКТ', 'НОЯ', 'ДЕК']
    plan_y = np.array(equipment.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(equipment.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([equipment.iloc[2, 11]])
    make_graph(ax[0, 1], "ОБОРУДОВАНИЕ", x, plan_y, fact_y, pred_y, month)
    
    plan_y = np.array(smr.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(smr.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([smr.iloc[2, 11]])
    make_graph(ax[1, 0], "СМР", x, plan_y, fact_y, pred_y, month)
    
    plan_y = np.array(other.iloc[0])
    fact_y = list([])
    i = 0
    while i != general['date'][0].month - 1:
        fact_y.append(other.iloc[1, i])
        i += 1
    fact_y = np.array(fact_y)
    pred_y = np.array([other.iloc[2, 11]])
    make_graph(ax[1, 1], "ПРОЧЕЕ", x, plan_y, fact_y, pred_y, month)
    
    plt.savefig("system_photo/dev_plan_structure.png", dpi=300, bbox_inches='tight')