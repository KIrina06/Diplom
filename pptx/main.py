### импорт библиотек
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date, datetime
import pandas as pd
import os
from io import BytesIO

st.set_page_config(layout="wide")

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

### импорт из файлов
from first_lvl_graph import draw_graph_1_lvl
from key_events_graph import draw_key_events_graph, draw_key_events
from dev_plan_graph_diag import draw_dev_plan_graph_accum, draw_dev_plan_diag_1, draw_dev_plan_diag_2
from dev_plan_structure_graph import draw_dev_plan_structure
from real_plan_graph import draw_real_plan_graph_accum
from foreign_revenue_graph import draw_foreign_revenue_graph_accum
from RD_SMR_graphs import draw_RD_SMR_graphs
from big_table import draw_big_table
from inventory_diagram import draw_inventory_diagram

from ppt_func import resize_text, name_of_slide, draw_table, head, head2

def main():
    st.title('Конструктор для создания презентации')
    st.write(f'Разработано: частным учереждением Госкорпорации "Росатом" "Отраслевой центр капитального строительства" Распространяется под Apache License 2.0 от 01.2004 http://www.apache.org/licenses/// created 2010.09.22 Разработано Кулешовой И. А. Создано 2024-07-12')
     
    #uploaded_file = st.file_uploader("Выберите файл", type=['xlsx'])
    uploaded_file = 'form.xlsx'
    if uploaded_file is None:
        st.write("Выберите, пожалуйста, файл!")
    else:
        #st.write("Файл успешно загружен!")
    
        # Загрузка, редактирование и обработка данных
        file = uploaded_file
        xl = pd.ExcelFile(file)
        
        general = xl.parse('general')
        
        program_execution = xl.parse('program_execution')

        development_plan = xl.parse('development_plan')

        for i in range(6):
            if development_plan.iloc[i, 4] != 0:
                development_plan.iloc[i, 4] = development_plan.iloc[i, 4][2 : (len(development_plan.iloc[i, 4]) - 1)]
                development_plan.iloc[i, 4] = development_plan.iloc[i, 4].replace(",", ".")
            development_plan.iloc[i, 4] = float(development_plan.iloc[i, 4])

        for i in range(6):
            if development_plan.iloc[i, 6] != '-':
                development_plan.iloc[i, 6] = str(development_plan.iloc[i, 6])[1 : (len(development_plan.iloc[i, 6]) - 2)]
            else: 
                development_plan.iloc[i, 6] = development_plan.iloc[i, 6].replace("-", "0")
            
            development_plan.iloc[i, 6] = int(development_plan.iloc[i, 6])

        for i in range(6):
            development_plan.iloc[i, 0] *= 100
            if development_plan.iloc[i, 5] != '-':
                development_plan.iloc[i, 5] *= 100
            else:
                development_plan.iloc[i, 5] = development_plan.iloc[i, 5].replace("-", "0")
            development_plan.iloc[i, 7] *= 100
            development_plan.iloc[i, 7] = int(development_plan.iloc[i, 7] // 1)
            development_plan.iloc[i, 8] *= 100

        values = [0 for col in development_plan.columns]
        for i in range(len(development_plan.iloc[0])):
            if i == 0 or i == 5 or i == 6 or i == 7 or i == 8:
                values[i] = int((development_plan.iloc[5, i] + development_plan.iloc[0, i]) / 2)
            else:
                values[i] = development_plan.iloc[5, i] + development_plan.iloc[0, i]
        development_plan2 = pd.DataFrame(np.insert(development_plan.values, 6, values, axis= 0 ))
        development_plan2.columns = development_plan.columns

        development_plan = pd.DataFrame(development_plan2)
        development_plan.columns = development_plan2.columns

        for i in range(6):
            development_plan.iloc[i, 7] = int(development_plan.iloc[i, 7])
            development_plan.iloc[i, 5] = int(development_plan.iloc[i, 5])

        realization_plan = xl.parse('realization_plan')
        for i in range(6):
            if realization_plan.iloc[i, 4] != 0:
                realization_plan.iloc[i, 4] = realization_plan.iloc[i, 4][2 : (len(realization_plan.iloc[i, 4]) - 1)]
                realization_plan.iloc[i, 4] = realization_plan.iloc[i, 4].replace(",", ".")
            realization_plan.iloc[i, 4] = float(realization_plan.iloc[i, 4])

        for i in range(6):
            if realization_plan.iloc[i, 6] != '-':
                realization_plan.iloc[i, 6] = str(realization_plan.iloc[i, 6])[1 : (len(realization_plan.iloc[i, 6]) - 2)]
            else: 
                realization_plan.iloc[i, 6] = realization_plan.iloc[i, 6].replace("-", "0")
            
            realization_plan.iloc[i, 6] = int(realization_plan.iloc[i, 6])

        for i in range(6):
            realization_plan.iloc[i, 0] *= 100
            if realization_plan.iloc[i, 5] != '-':
                realization_plan.iloc[i, 5] *= 100
            else:
                realization_plan.iloc[i, 5] = realization_plan.iloc[i, 5].replace("-", "0")
            realization_plan.iloc[i, 7] *= 100
            realization_plan.iloc[i, 7] = int(realization_plan.iloc[i, 7] // 1)
            realization_plan.iloc[i, 8] *= 100

        values = [0 for col in realization_plan.columns]
        for i in range(len(realization_plan.iloc[0])):
            if i == 0 or i == 5 or i == 6 or i == 7 or i == 8:
                values[i] = int((realization_plan.iloc[5, i] + realization_plan.iloc[0, i]) / 2)
            else:
                values[i] = realization_plan.iloc[5, i] + realization_plan.iloc[0, i]
        realization_plan2 = pd.DataFrame(np.insert(realization_plan.values, 6, values, axis= 0 ))
        realization_plan2.columns = realization_plan.columns

        realization_plan = pd.DataFrame(realization_plan2)
        realization_plan.columns = realization_plan2.columns

        for i in range(6):
            realization_plan.iloc[i, 7] = int(realization_plan.iloc[i, 7])
            realization_plan.iloc[i, 5] = int(realization_plan.iloc[i, 5])

        accumulative_execution = xl.parse('accumulative_execution')
        foreign_revenue = xl.parse('foreign_revenue')
        for i in range(4):
            if foreign_revenue.iloc[i, 3] != 0:
                foreign_revenue.iloc[i, 3] = foreign_revenue.iloc[i, 3][2 : (len(foreign_revenue.iloc[i, 3]) - 1)]
                foreign_revenue.iloc[i, 3] = foreign_revenue.iloc[i, 3].replace(",", ".")
            foreign_revenue.iloc[i, 3] = float(foreign_revenue.iloc[i, 3])

        for i in range(4):
            if foreign_revenue.iloc[i, 5] != '-':
                foreign_revenue.iloc[i, 5] = str(foreign_revenue.iloc[i, 5])[1 : (len(foreign_revenue.iloc[i, 5]) - 2)]
            else: 
                foreign_revenue.iloc[i, 5] = foreign_revenue.iloc[i, 5].replace("-", "0")
            
            foreign_revenue.iloc[i, 5] = int(foreign_revenue.iloc[i, 5])

        for i in range(4):
            if foreign_revenue.iloc[i, 4] != '-':
                foreign_revenue.iloc[i, 4] *= 100
            else:
                foreign_revenue.iloc[i, 4] = foreign_revenue.iloc[i, 4].replace("-", "0")
            foreign_revenue.iloc[i, 6] *= 100
            foreign_revenue.iloc[i, 6] = int(foreign_revenue.iloc[i, 6] // 1)
            foreign_revenue.iloc[i, 7] *= 100

        values = [0 for col in foreign_revenue.columns]
        for i in range(len(foreign_revenue.iloc[0])):
            if i == 4 or i == 5 or i == 6 or i == 7:
                values[i] = int((foreign_revenue.iloc[0, i] + foreign_revenue.iloc[1, i] + foreign_revenue.iloc[2, i] + foreign_revenue.iloc[3, i]) / 4)
            else:
                values[i] = foreign_revenue.iloc[0, i] + foreign_revenue.iloc[1, i] + foreign_revenue.iloc[2, i] + foreign_revenue.iloc[3, i]
        foreign_revenue2 = pd.DataFrame(np.insert(foreign_revenue.values, 4, values, axis= 0 ))
        foreign_revenue2.columns = foreign_revenue.columns

        foreign_revenue = pd.DataFrame(foreign_revenue2)
        foreign_revenue.columns = foreign_revenue2.columns
        
        risk_assessment = xl.parse('risk_assessment')
        info = xl.parse('Info')
        status = xl.parse('Status')
        
        # Титульный лист
        st.subheader("Титульный лист")
        
        # Загрузка, редактирование и обработка данных
        text = st.text_input("Название сооружения: ", value = general['object_name'][0])
        if text != "":
            general['object_name'][0] = text

        types = ["сооружение", "корабль"]
        text = st.selectbox("Тип сооружения: ", types)
        if text != "":
            general['object_type'][0] = text
        
        month = general['date'][0].month
        if month // 10 == 0: 
            month = f'0{month}'

        day = general['date'][0].day
        if day // 10 == 0: 
            day = f'0{day}'
        
        date_str = st.text_input("Дата (для рассмотрения на Операционном комитете) (ДД.ММ.ГГГГ):", value = str(f"{day}.{month}.{general['date'][0].year}"))
        if date_str != str(f"{day}.{month}.{general['date'][0].year}"):
            # --- Преобразуем строку в объект datetime ---
            try:
                selected_date = datetime.strptime(date_str, "%d.%m.%Y").date()
                general['date'][0] = pd.Timestamp(selected_date)
            except ValueError:
                st.error("Неверный формат даты. Используйте ДД.ММ.ГГГГ.")
        
        month = general['date_now'][0].month
        if month // 10 == 0: 
            month = f'0{month}'

        day = general['date_now'][0].day
        if day // 10 == 0: 
            day = f'0{day}'
        
        date_str = st.text_input("Текущая дата:", value = str(f"{day}.{month}.{general['date_now'][0].year}"))

        if date_str != str(f"{day}.{month}.{general['date_now'][0].year}"):
            # --- Преобразуем строку в объект datetime ---
            try:
                selected_date = datetime.strptime(date_str, "%d.%m.%Y").date()
                general['date_now'][0] = pd.Timestamp(selected_date)
            except ValueError:
                st.error("Неверный формат даты. Используйте ДД.ММ.ГГГГ.")

        text = st.text_input("ФИО докладчика: ", value = general['fio'][0])
        if text != "":
            general['fio'][0] = text

        text = st.text_input("Должность докладчика: ", value = general['job'][0])
        if text != "":
            general['job'][0] = text
        
        # График 1-ого уровня
        st.subheader("График 1-ого уровня")
        
        # Загрузка, редактирование и обработка данных
        graph_1_lvl = xl.parse('graph_1_lvl')
        edited_graph_1_lvl = pd.DataFrame(graph_1_lvl)
        edited_graph_1_lvl.columns = ["Название события", "Дата выполнения,\nутвержденная по графику", "Фактическая/по прогнозу\n дата выполнения"]
        edited_graph_1_lvl = st.data_editor(edited_graph_1_lvl, num_rows="dynamic")
        
        col_names_graph_1_lvl = graph_1_lvl.columns.tolist()
        
        graph_1_lvl = pd.DataFrame(edited_graph_1_lvl)
        graph_1_lvl.columns = col_names_graph_1_lvl
        
        # Создаем графики
        draw_graph_1_lvl(general, graph_1_lvl)
        st.image("system_photo/1lvlgraph.png")
        
        # Ключевые события
        st.subheader("Ключевые события")
        
        # Загрузка, редактирование и обработка данных
        key_events = xl.parse('key_events')
        edited_key_events = pd.DataFrame(key_events)
        edited_key_events.columns = ["Наименование ключевого события", "Признак", "Объем по ключевому событию", "Факт. выпол.", "% вып.", 
        "Дата выполнения, утвержденная планом", "Прогнозируемая дата выполнения", "НС, влияющие на ключевые события"]
        edited_key_events = st.data_editor(edited_key_events, num_rows="dynamic")
        
        col_names_key_events = key_events.columns.tolist()
        
        key_events = pd.DataFrame(edited_key_events)
        key_events.columns = col_names_key_events
        
        # Создаем графики
        draw_key_events_graph(general, key_events)
        files = os.listdir("system_photo")
        for file in files:
            if "key_events_" in file:
                st.image("system_photo/" + file)
                
        # Выполнение плана по освоению
        st.subheader("Выполнение плана по освоению")
        
        # Загрузка, редактирование и обработка данных
        
        edited_accumulative_execution = pd.DataFrame(accumulative_execution)
        edited_accumulative_execution.index = ["План", "Факт", "Прогноз"]
        edited_accumulative_execution.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_accumulative_execution = st.data_editor(edited_accumulative_execution, num_rows="dynamic")
        
        col_names_accumulative_execution = accumulative_execution.columns.tolist()
        
        accumulative_execution = pd.DataFrame(edited_accumulative_execution)
        accumulative_execution.columns = col_names_accumulative_execution
        
        # Создаем графики
        draw_dev_plan_graph_accum(general, accumulative_execution)
        st.image("system_photo/dev_plan_graph_accum.png")
        
        # Загрузка, редактирование и обработка данных
        dev_plan_diag_1 = xl.parse('dev_plan_diag_1')
        edited_dev_plan_diag_1 = pd.DataFrame(dev_plan_diag_1)
        edited_dev_plan_diag_1.index = ["План", "Факт"]
        edited_dev_plan_diag_1.columns = ["ДСО (собственные силы)", "Сторонние подрядчики"]
        edited_dev_plan_diag_1 = st.data_editor(edited_dev_plan_diag_1, num_rows="dynamic")
        
        col_names_dev_plan_diag_1 = dev_plan_diag_1.columns.tolist()
        
        dev_plan_diag_1 = pd.DataFrame(edited_dev_plan_diag_1)
        dev_plan_diag_1.columns = col_names_dev_plan_diag_1
        
        # Создаем графики
        draw_dev_plan_diag_1(dev_plan_diag_1)
        st.image("system_photo/dev_plan_diag_1.png")
        
        # Загрузка, редактирование и обработка данных
        dev_plan_diag_2 = xl.parse('dev_plan_diag_2')
        edited_dev_plan_diag_2 = pd.DataFrame(dev_plan_diag_2)
        edited_dev_plan_diag_2.index = ["План", "Факт"]
        edited_dev_plan_diag_2.columns = ["% от общей годовой программы СМР", "количествово комплектов РД"]
        edited_dev_plan_diag_2 = st.data_editor(edited_dev_plan_diag_2, num_rows="dynamic")
        
        col_names_dev_plan_diag_2 = dev_plan_diag_2.columns.tolist()
        
        dev_plan_diag_2 = pd.DataFrame(edited_dev_plan_diag_2)
        dev_plan_diag_2.columns = col_names_dev_plan_diag_2
        
        # Создаем графики
        draw_dev_plan_diag_2(dev_plan_diag_2)
        st.image("system_photo/dev_plan_diag_2.png")
        
        # Выполнение плана по освоению по структуре затрат
        st.subheader("Выполнение плана по освоению по структуре затрат")
        
        # Загрузка, редактирование и обработка данных
        pir = xl.parse('pir')
        equipment = xl.parse('equipment')
        smr = xl.parse('smr')
        other = xl.parse('other')
        
        # ПИР
        st.write("ПИР")
        edited_pir = pd.DataFrame(pir)
        edited_pir.index = ["План", "Факт", "Прогноз"]
        edited_pir.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_pir = st.data_editor(edited_pir, num_rows="dynamic")
        
        col_names_pir = pir.columns.tolist()
        
        pir = pd.DataFrame(edited_pir)
        pir.columns = col_names_pir
        
        # ОБОРУДОВАНИЕ
        st.write("Оборудование")
        edited_equipment = pd.DataFrame(equipment)
        edited_equipment.index = ["План", "Факт", "Прогноз"]
        edited_equipment.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_equipment = st.data_editor(edited_equipment, num_rows="dynamic")
        
        col_names_equipment = equipment.columns.tolist()
        
        equipment = pd.DataFrame(edited_equipment)
        equipment.columns = col_names_equipment
        
        # СМР
        st.write("СМР")
        edited_smr = pd.DataFrame(smr)
        edited_smr.index = ["План", "Факт", "Прогноз"]
        edited_smr.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_smr = st.data_editor(edited_smr, num_rows="dynamic")
        
        col_names_smr = smr.columns.tolist()
        
        smr = pd.DataFrame(edited_smr)
        smr.columns = col_names_smr
        
        # ДРУГОЕ
        st.write("Другое")
        edited_other = pd.DataFrame(other)
        edited_other.index = ["План", "Факт", "Прогноз"]
        edited_other.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_other = st.data_editor(edited_other, num_rows="dynamic")
        
        col_names_other = other.columns.tolist()
        
        other = pd.DataFrame(edited_other)
        other.columns = col_names_other
        
        # Создаем графики
        draw_dev_plan_structure(general, pir, equipment, smr, other)
        st.image("system_photo/dev_plan_structure.png")
        
        # Выполнение плана по реализации
        st.subheader("Выполнение плана по реализации")
        
        # Загрузка, редактирование и обработка данных
        accumulative_realization = xl.parse('accumulative_realization')
        edited_accumulative_realization = pd.DataFrame(accumulative_realization)
        edited_accumulative_realization.index = ["План", "Факт", "Прогноз"]
        edited_accumulative_realization.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_accumulative_realization = st.data_editor(edited_accumulative_realization, num_rows="dynamic")
        
        col_names_accumulative_realization = accumulative_realization.columns.tolist()
        
        accumulative_realization = pd.DataFrame(edited_accumulative_realization)
        accumulative_realization.columns = col_names_accumulative_realization
        
        # Создаем графики
        draw_real_plan_graph_accum(general, accumulative_realization)
        st.image("system_photo/real_plan_graph.png")
        
        # Выполнение показателя "зарубежная выручка"
        st.subheader('Выполнение показателя "зарубежная выручка"')
        
        # Загрузка, редактирование и обработка данных
        foreign_revenue_accum = xl.parse('foreign_revenue_accum')
        edited_foreign_revenue_accum = pd.DataFrame(foreign_revenue_accum)
        edited_foreign_revenue_accum.index = ["План", "Факт", "Прогноз"]
        edited_foreign_revenue_accum.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_foreign_revenue_accum = st.data_editor(edited_foreign_revenue_accum, num_rows="dynamic")
        
        col_names_foreign_revenue_accum = foreign_revenue_accum.columns.tolist()
        
        foreign_revenue_accum = pd.DataFrame(edited_foreign_revenue_accum)
        foreign_revenue_accum.columns = col_names_foreign_revenue_accum
        
        # Создаем графики
        draw_foreign_revenue_graph_accum(general, foreign_revenue_accum)
        st.image("system_photo/foreign_revenue_graph.png")
        
        # Статус выдачи РД на объем СМР
        st.subheader("Статус выдачи РД на объем СМР")
        
        # Загрузка, редактирование и обработка данных
        RD_month = xl.parse('RD_month')
        # Выдача корректировки РД, необходимой для СМР, ед. (ПО МЕСЯЦАМ)
        st.write("Выдача корректировки РД, необходимой для СМР, ед. (ПО МЕСЯЦАМ)")
        edited_RD_month = pd.DataFrame(RD_month)
        edited_RD_month.index = ["План", "Факт", "Прогноз"]
        edited_RD_month.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_RD_month = st.data_editor(edited_RD_month, num_rows="dynamic")
        
        col_names_RD_month = RD_month.columns.tolist()
        
        RD_month = pd.DataFrame(edited_RD_month)
        RD_month.columns = col_names_RD_month
        
        values = [RD_month.iloc[1][col] - RD_month.iloc[0][col] for col in RD_month.columns]
        RD_month2 = pd.DataFrame(np.insert(RD_month.values, 3, values, axis= 0 ))
        RD_month2.columns = RD_month.columns

        months = ['jan', 'feb',	'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        # получаем текущую дату
        month = general['date'][0].month
        for i in range(len(RD_month2.iloc[3])):
            if months.index(RD_month2.columns[i]) >= month - 1:
                RD_month2.iloc[3, i] = 0

        values = []
        for col in RD_month2.columns:
            if RD_month2.iloc[0][col] != 0:
                values.append(round((RD_month2.iloc[1][col] / RD_month2.iloc[0][col] * 100), 0))
            else:
                values.append(0)
        RD_month = pd.DataFrame(np.insert(RD_month2.values, 4, values, axis= 0 ))
        RD_month.columns = RD_month2.columns
        
        # Выдача корректировки РД, необходимой для СМР, ед. (НАКОПИТЕЛЬНО)
        RD_accumulative = xl.parse('RD_accumulative')
        
        st.write("Выдача корректировки РД, необходимой для СМР, ед. (НАКОПИТЕЛЬНО)")
        
        edited_RD_accumulative = pd.DataFrame(RD_accumulative)
        edited_RD_accumulative.index = ["План", "Факт", "Прогноз"]
        edited_RD_accumulative.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_RD_accumulative = st.data_editor(edited_RD_accumulative, num_rows="dynamic")
        
        col_names_RD_accumulative = RD_accumulative.columns.tolist()
        
        RD_accumulative = pd.DataFrame(edited_RD_accumulative)
        RD_accumulative.columns = col_names_RD_accumulative
        
        values = [RD_accumulative.iloc[1][col] - RD_accumulative.iloc[0][col] for col in RD_accumulative.columns]
        RD_accumulative2 = pd.DataFrame(np.insert(RD_accumulative.values, 3, values, axis= 0 ))
        RD_accumulative2.columns = RD_accumulative.columns

        months = ['jan', 'feb',	'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        # получаем текущую дату
        month = general['date'][0].month
        for i in range(len(RD_accumulative2.iloc[3])):
            if months.index(RD_accumulative2.columns[i]) >= month - 1:
                RD_accumulative2.iloc[3, i] = 0

        values = []
        for col in RD_accumulative2.columns:
            if RD_accumulative2.iloc[0][col] != 0:
                values.append(round((RD_accumulative2.iloc[1][col] / RD_accumulative2.iloc[0][col] * 100), 0))
            else:
                values.append(0)

        RD_accumulative = pd.DataFrame(np.insert(RD_accumulative2.values, 4, values, axis= 0 ))
        RD_accumulative.columns = RD_accumulative2.columns
        
        # Создаем графики
        draw_RD_SMR_graphs(general, RD_month, RD_accumulative)
        st.image("system_photo/RD_SMR_graphs.png")
        
        # Численность строительного персонала на площадке
        st.subheader("Численность строительного персонала на площадке")
        
        # Загрузка, редактирование и обработка данных
        num_of_builders = xl.parse('num_of_builders')

        edited_num_of_builders = pd.DataFrame(num_of_builders)
        edited_num_of_builders.index = ["ДСО (собственные силы): План", "ДСО (собственные силы): в т. ч. ИТР", "ДСО (собственные силы): Факт/Прогноз",
                "ДСО (собственные силы): в т. ч. ИТР", "Сторонние организации: План", "Сторонние организации: в т. ч. ИТР", "Сторонние организации: Факт/Прогноз", 
                "Сторонние организации: в т. ч. ИТР"]
        edited_num_of_builders.columns = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
        edited_num_of_builders = st.data_editor(edited_num_of_builders, num_rows="dynamic")
        
        col_names_num_of_builders = num_of_builders.columns.tolist()
        
        num_of_builders = pd.DataFrame(edited_num_of_builders)
        num_of_builders.columns = col_names_num_of_builders

        values = [num_of_builders.iloc[2][col] - num_of_builders.iloc[0][col] for col in num_of_builders.columns]
        num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 4, values, axis= 0 ))
        num_of_builders2.columns = num_of_builders.columns

        values = [num_of_builders2.iloc[7][col] - num_of_builders2.iloc[5][col] for col in num_of_builders2.columns]
        num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 9, values, axis= 0 ))
        num_of_builders.columns = num_of_builders2.columns

        values = [num_of_builders.iloc[0][col] + num_of_builders.iloc[5][col] for col in num_of_builders.columns]
        num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 10, values, axis= 0 ))
        num_of_builders2.columns = num_of_builders.columns

        values = [num_of_builders2.iloc[1][col] + num_of_builders2.iloc[6][col] for col in num_of_builders2.columns]
        num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 11, values, axis= 0 ))
        num_of_builders.columns = num_of_builders2.columns

        values = [num_of_builders.iloc[2][col] + num_of_builders.iloc[7][col] for col in num_of_builders.columns]
        num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 12, values, axis= 0 ))
        num_of_builders2.columns = num_of_builders.columns

        values = [num_of_builders2.iloc[3][col] + num_of_builders2.iloc[8][col] for col in num_of_builders2.columns]
        num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 13, values, axis= 0 ))
        num_of_builders.columns = num_of_builders2.columns

        values = [num_of_builders.iloc[12][col] - num_of_builders.iloc[10][col] for col in num_of_builders.columns]
        num_of_builders2 = pd.DataFrame(np.insert(num_of_builders.values, 14, values, axis= 0 ))
        num_of_builders2.columns = num_of_builders.columns

        values = [round((num_of_builders2.iloc[12][col] / num_of_builders2.iloc[10][col] * 100), 0) for col in num_of_builders2.columns]
        num_of_builders = pd.DataFrame(np.insert(num_of_builders2.values, 15, values, axis= 0 ))
        num_of_builders.columns = num_of_builders2.columns
        
        # Создаем графики
        draw_big_table(num_of_builders)
        st.image("system_photo/big_table.png")
        
        # Инвентаризация проектов на предмет наличия рисков срыва поставок оборудования и комплектующих из 3-х стран
        st.subheader("Инвентаризация проектов на предмет наличия рисков срыва поставок оборудования и комплектующих из 3-х стран")
        
        # Загрузка, редактирование и обработка данных
        inventory = xl.parse('inventory')

        edited_inventory = pd.DataFrame(inventory)
        edited_inventory.columns = ["Законтрактовано, млн долл.: Риск отсутствует", "Законтрактовано, млн долл.: Риск срыва поставок ввиду санкционного давления на РФ", 
                "Не законтрактовано, млн долл.: Риск отсутствует", "Не законтрактовано, млн долл.: Риск срыва поставок  ввиду санкционного давления на РФ"]
        edited_inventory = st.data_editor(edited_inventory, num_rows="dynamic")
        
        col_names_inventory = inventory.columns.tolist()
        
        inventory = pd.DataFrame(edited_inventory)
        inventory.columns = col_names_inventory
        
        # Создаем графики
        draw_inventory_diagram(inventory)
        st.image("system_photo/inventory_diagram.png")
        
        xl.close()
        
        button = st.button('Сгенерировать презентацию')

        if button:
            ppt = Presentation()  
            # задаем параметры слайдов (высота и ширина)
            ppt.slide_height = Inches(9) 
            ppt.slide_width = Inches(16)
            # за основу для слайдов берем пустой шаблон слайда
            blank_slide_layout = ppt.slide_layouts[6]  
            
            # Титульный лист
            slide = ppt.slides.add_slide(blank_slide_layout)
            # задаем фон
            background = slide.shapes.add_picture("system_photo/background.png", Cm(0), Cm(0), width=Inches(16), height=Inches(9))
            # рисуем иконку
            pic = slide.shapes.add_picture("system_photo/icon.png", Cm(1.38), Cm(1.87), width=Cm(5.88), height=Cm(5.11))

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
            
            # График 1-ого уровня
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = "НА ВЕСЬ ПЕРИОД СТРОИТЕЛЬСТВА"
            name_of_slide(slide, f'ГРАФИК 1-ГО УРОВНЯ {object_name}', subinf)
            
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
                            
            # Ключевые события
            key_len = len(key_events['event_name'])
            for file in files:
                if "key_events_" in file:
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    # добавление названия слайда
                    subinf = "С УКАЗАНИЕМ ФИЗИЧЕСКИХ ОБЪЕМОВ РАБОТ"
                    name_of_slide(slide, f'КЛЮЧЕВЫЕ СОБЫТИЯ {general["date"][0].year} ГОДА', subinf)
                    if key_len >= 4:
                        pic = slide.shapes.add_picture("system_photo/" + file, Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
                        key_len -= 4                
                    else:
                        pic = slide.shapes.add_picture("system_photo/" + file, Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4 / (4 + 2 / 3) * (key_len + 2 / 3)))
                    pic = slide.shapes.add_picture("system_photo/legend2.png", Cm(1.25), Cm(22.11), width=Cm(15.06), height=Cm(0.73))
            
            # Выполнение плана по освоению
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

            for i in range(3, 10):
                for j in range(2, 11):
                    if j == 2 or j == 7 or j == 9 or j == 10:
                        table.cell(i, j).text = str(development_plan.iloc[i - 3, j - 2]) + "%"
                        if j == 7 and development_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"
                    elif j == 3 or j == 4 or j == 5:
                        table.cell(i, j).text = str(development_plan.iloc[i - 3, j - 2])
                    elif j == 6:
                        table.cell(i, j).text = "(+" + str(development_plan.iloc[i - 3, j - 2]) + ")"
                        if development_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "0"
                    else:
                        table.cell(i, j).text = "(" + str(development_plan.iloc[i - 3, j - 2]) + "%)"
                        if j == 8 and development_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"

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
            
            pic = slide.shapes.add_picture("system_photo/dev_plan_graph_accum.png", Inches(0.161417), Inches(4.77), width=Inches(11.46), height=Inches(3.85))
            pic = slide.shapes.add_picture("system_photo/dev_plan_diag_1.png", Cm(29.3), Cm(11.74), width=Cm(10.83), height=Cm(6.8))
            pic = slide.shapes.add_picture("system_photo/dev_plan_diag_2.png", Cm(31.67), Cm(17.87), width=Cm(7.49), height=Cm(4.23))
            
            # Выполнение плана по освоению по структуре затрат
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = f'ПО ПРОЕКТУ {object_name}'
            name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПЛАНА ПО ОСВОЕНИЮ {general["date"][0].year} ГОДА ПО СТРУКТУРЕ ЗАТРАТ', subinf)
            
            draw_dev_plan_structure(general, pir, equipment, smr, other)
            
            pic = slide.shapes.add_picture("system_photo/dev_plan_structure.png", Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
            
            # Выполнение плана по реализации
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

            for i in range(3, 10):
                for j in range(2, 11):
                    if j == 2 or j == 7 or j == 9 or j == 10:
                        table.cell(i, j).text = str(realization_plan.iloc[i - 3, j - 2]) + "%"
                        if j == 7 and realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"
                    elif j == 3 or j == 4 or j == 5:
                        table.cell(i, j).text = str(realization_plan.iloc[i - 3, j - 2])
                    elif j == 6:
                        table.cell(i, j).text = "(+" + str(realization_plan.iloc[i - 3, j - 2]) + ")"
                        if realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "0"
                    else:
                        table.cell(i, j).text = "(" + str(realization_plan.iloc[i - 3, j - 2]) + "%)"
                        if j == 8 and realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"

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
            
            pic = slide.shapes.add_picture("system_photo/real_plan_graph.png", Cm(2.71), top = Cm(11.53), width=Cm(34.96), height=Cm(10.72))
            
            # Выполнение показателя "зарубежная выручка"
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = f'ПО ПРОЕКТУ {object_name}'
            name_of_slide(slide, f'ВЫПОЛНЕНИЕ ПОКАЗАТЕЛЯ "ЗАРУБЕЖНАЯ ВЫРУЧКА" В {general["date"][0].year} ГОДУ', subinf)
            
            left = Cm(0.29)
            top = Cm(1.76)
            width = Cm(39.84)
            height = Cm(2.48)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            # Enable word wrap
            tf.word_wrap = True
            p = tf.add_paragraph()
            p.text = "ВЫПОЛНЕНИЕ ПРОГРАММЫ 2024 ГОДА (ЗАРУБЕЖНАЯ ВЫРУЧКА), млн. долл."
            # Optionally, if you want to center the entire textbox including its vertical position
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            # Center the paragraph text
            p.alignment = PP_ALIGN.LEFT
            p.font.size = Pt(20)
            p.font.name = "Arial Narrow"
            p.font.color.rgb = RGBColor(32, 56, 100)
            
            top = Cm(4.05)
            table = draw_table(shapes, top)

            for i in range(3, 10):
                for j in range(2, 11):
                    if j == 2 or j == 7 or j == 9 or j == 10:
                        table.cell(i, j).text = str(realization_plan.iloc[i - 3, j - 2]) + "%"
                        if j == 7 and realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"
                    elif j == 3 or j == 4 or j == 5:
                        table.cell(i, j).text = str(realization_plan.iloc[i - 3, j - 2])
                    elif j == 6:
                        table.cell(i, j).text = "(+" + str(realization_plan.iloc[i - 3, j - 2]) + ")"
                        if realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "0"
                    else:
                        table.cell(i, j).text = "(" + str(realization_plan.iloc[i - 3, j - 2]) + "%)"
                        if j == 8 and realization_plan.iloc[i - 3, j - 2] == 0:
                            table.cell(i, j).text = "-"

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
            
            pic = slide.shapes.add_picture("system_photo/foreign_revenue_graph.png", Cm(2.71), top = Cm(11.53), width=Cm(34.96), height=Cm(10.72))
            
            # Оценка рисков
            ind = 0
            ost_rows = 23
            if len(risk_assessment['risk_name']) == 0:
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'ОЦЕНКА РИСКОВ И ПЛАН ПРЕДЛАГАЕМЫХ КОМПЕНСИРУЮЩИХ МЕРОПРИЯТИЙ', subinf)
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(10.72)
                    top = Cm(11.38)
                    width = Cm(19.2)
                    height = Cm(2.3)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    
                    # разрешаем перенос слов
                    tf.word_wrap = True
                        
                    # добавляем параграф текста
                    p = tf.add_paragraph()
                    p.text = "В рамках проекта сооружения риски\nза отчетный период отсутствуют"
                    # выравниваем по центрк по вертикали
                    tf.vertical_anchor = MSO_ANCHOR.TOP
                    # выравниваем по центру по горизонтали
                    p.alignment = PP_ALIGN.CENTER
                    p.font.size = Pt(18)
                    p.font.name = "Arial Narrow"
                    p.font.color.rgb = RGBColor(32, 56, 100)
                    
            for i in range(len(risk_assessment['risk_name'])):
                num_rows_risk_name = len(risk_assessment['risk_name'][i]) / 27
                if num_rows_risk_name > len(risk_assessment['risk_name'][i]) // 27:
                    num_rows_risk_name = int(len(risk_assessment['risk_name'][i]) // 27 + 1)
                num_rows_measures = len(str(risk_assessment['measures'])[i]) / 24
                if num_rows_measures > len(str(risk_assessment['measures'])[i]) // 24:
                    num_rows_measures = int(len(str(risk_assessment['measures'])[i]) // 24 + 1)
                num_rows_status = len(str(risk_assessment['status'])[i]) / 24
                if num_rows_status > len(str(risk_assessment['status'])[i]) // 24:
                    num_rows_status = int(len(str(risk_assessment['status'])[i]) // 24 + 1)
                num_rows_cell = 0

                x = np.array([num_rows_risk_name, num_rows_measures, num_rows_status])
                num_rows_cell = x.max()
                ost_rows = ost_rows - num_rows_cell - 3

                if ost_rows >= 0 and i == len(risk_assessment['risk_name']) - 1:
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'ОЦЕНКА РИСКОВ И ПЛАН ПРЕДЛАГАЕМЫХ КОМПЕНСИРУЮЩИХ МЕРОПРИЯТИЙ', subinf)
                    
                    head(shapes)
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(2.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(13.63)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{risk_assessment['risk_name'][j]}" + "\n" * int(num_rows_cell - num_rows_risk_name + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(0.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(2)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f'{j + 1}' + "\n" * int(num_rows_cell + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(16.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(12)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if risk_assessment['measures'][j] == None or risk_assessment['measures'][j] == "nan":
                            risk_assessment['measures'][j] = " "
                        p.text = f"{risk_assessment['measures'][j]}" + "\n" * int(num_rows_cell - num_rows_measures + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(28.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(12)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if risk_assessment['status'][j] == None or risk_assessment['status'][j] == "nan":
                            risk_assessment['status'][j] = " "
                        p.text = f"{risk_assessment['status'][j]}" + "\n" * int(num_rows_cell - num_rows_status + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                        
                elif ost_rows < 0:
                    ost_rows = 23
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'ОЦЕНКА РИСКОВ И ПЛАН ПРЕДЛАГАЕМЫХ КОМПЕНСИРУЮЩИХ МЕРОПРИЯТИЙ', subinf)
                    
                    head(shapes)
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(2.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(13.63)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{risk_assessment['risk_name'][j]}" + "\n" * int(num_rows_cell - num_rows_risk_name + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(0.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(2)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f'{j + 1}' + "\n" * int(num_rows_cell + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(16.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(12)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if risk_assessment['measures'][j] == None or risk_assessment['measures'][j] == "nan":
                            risk_assessment['measures'][j] = " "
                        p.text = f"{risk_assessment['measures'][j]}" + "\n" * int(num_rows_cell - num_rows_measures)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(28.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(12)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if risk_assessment['status'][j] == None or risk_assessment['status'][j] == "nan":
                            risk_assessment['status'][j] = " "
                        p.text = f"{risk_assessment['status'][j]}" + "\n" * int(num_rows_cell - num_rows_status + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    ind = i
            
            # Приложения
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            app_flag = len(ppt.slides) - 1
            
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
            
            # Статус выдачи РД на объем СМР
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = ""
            name_of_slide(slide, f'СТАТУС ВЫДАЧИ РД НА ОБЪЕМ СМР {general["date"][0].year} ГОДУ', subinf)
            
            pic = slide.shapes.add_picture("system_photo/RD_SMR_graphs.png", Inches(0.401575), Cm(3.12), width=Inches(15.2), height=Inches(7.4))
            
            # Численность строительного персонала на площадке
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = ""
            name_of_slide(slide, f'ЧИСЛЕННОСТЬ СТРОИТЕЛЬНОГО ПЕРСОНАЛА НА ПЛОЩАДКЕ В {general["date"][0].year} ГОДУ', subinf)
            
            pic = slide.shapes.add_picture("system_photo/big_table.png", Inches(0.401575), Inches(1.4), width=Inches(15.2), height=Inches(7.4))
            
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
                    
            # Инвентаризация проектов на предмет наличия рисков срыва поставок оборудования и комплектующих из 3-х стран
            # создаем слайд
            slide = ppt.slides.add_slide(blank_slide_layout)
            shapes = slide.shapes
            subinf = ""
            name_of_slide(slide, f'ИНВЕНТАРИЗАЦИЯ ПРОЕКТА СООРУЖЕНИЕ {object_name} НА ПРЕДМЕТ НАЛИЧИЯ РИСКОВ СРЫВА ПОСТАВОК ОБОРУДОВАНИЯ И КОМПЛЕКТУЮЩИХ ИЗ 3-Х СТРАН', subinf)
            
            pic = slide.shapes.add_picture("system_photo/inventory_diagram.png", Inches(0.401575), Inches(1.2), width=Inches(15.2), height=Inches(7.4))
            
            # Справочно. Обеспеченность площадки строительства денежными средствами
            ind = 0
            ost_rows = 24
            for i in range(len(info['question'])):
                num_rows_question = len(info['question'][i]) / 58
                if num_rows_question > len(info['question'][i]) // 58:
                    num_rows_question = int(len(info['question'][i]) // 58 + 1)
                num_rows_answer = len(info['answer'][i]) / 78 + 1
                if num_rows_answer > len(info['answer'][i]) // 78 + 1:
                    num_rows_answer = int(len(info['answer'][i]) // 78 + 2)
                ost_rows = int(ost_rows - num_rows_question // 0.75 - num_rows_answer)
                if ost_rows >= 0 and i == len(info['question']) - 1:
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'СПРАВОЧНО. ОБЕСПЕЧЕННОСТЬ ПЛОЩАДКИ СТРОИТЕЛЬСТВА ДЕНЕЖНЫМИ СРЕДСТВАМИ', subinf)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Inches(0.5)
                    top = Cm(3.05)
                    width = Inches(16 - 2 * 0.5)
                    height = Cm(18.86)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{info['question'][j]}"
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        # делаем шрифт жирным, 40 размера, Arial Narrow, цвета (32, 56, 100)
                        p.font.bold = True
                        p.font.size = Pt(22)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                        p1 = tf.add_paragraph()
                        p1.text = f"{info['answer'][j]}\n"
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p1.alignment = PP_ALIGN.LEFT
                        p1.font.size = Pt(18)
                        p1.font.name = "Arial Narrow"
                        p1.font.color.rgb = RGBColor(32, 56, 100)
                elif ost_rows < 0:
                    ost_rows = 24
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'СПРАВОЧНО. ОБЕСПЕЧЕННОСТЬ ПЛОЩАДКИ СТРОИТЕЛЬСТВА ДЕНЕЖНЫМИ СРЕДСТВАМИ', subinf)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Inches(0.5)
                    top = Cm(3.05)
                    width = Inches(16 - 2 * 0.5)
                    height = Cm(18.86)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                        
                    # разрешаем перенос слов
                    tf.word_wrap = True
                    for j in range(ind, i):
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{info['question'][j]}"
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        # делаем шрифт жирным, 40 размера, Arial Narrow, цвета (32, 56, 100)
                        p.font.bold = True
                        p.font.size = Pt(22)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                        p1 = tf.add_paragraph()
                        p1.text = f"{info['answer'][j]}\n"
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p1.alignment = PP_ALIGN.LEFT
                        p1.font.size = Pt(18)
                        p1.font.name = "Arial Narrow"
                        p1.font.color.rgb = RGBColor(32, 56, 100)
                    ind = i
            
            # Статус выполнения поручений ген. директора
            ind = 0
            ost_rows = 23
            for i in range(len(status['name'])):
                num_rows_name = len(status['name'][i]) / 21
                if num_rows_name > len(status['name'][i]) // 21:
                    num_rows_name = int(len(status['name'][i]) // 21 + 1)
                num_rows_complete_date = len(str(status['complete_date'])[i]) / 13
                if num_rows_complete_date > len(str(status['complete_date'])[i]) // 13:
                    num_rows_complete_date = int(len(str(status['complete_date'])[i]) // 13 + 1)
                
                num_rows_people	 = len(str(status['people'])[i]) / 13
                if num_rows_people	 > len(str(status['people'])[i]) // 13:
                    num_rows_people	 = int(len(str(status['people'])[i]) // 13 + 1)

                num_rows_status	 = len(str(status['status'])[i]) / 11
                if num_rows_status	 > len(str(status['status'])[i]) // 11:
                    num_rows_status	 = int(len(str(status['status'])[i]) // 11 + 1)

                num_rows_comments	 = len(str(status['comments'])[i]) / 13
                if num_rows_comments	 > len(str(status['comments'])[i]) // 13:
                    num_rows_comments	 = int(len(str(status['comments'])[i]) // 13 + 1)
                num_rows_cell = 0

                x = np.array([num_rows_name, num_rows_complete_date, num_rows_people, num_rows_status, num_rows_comments])
                num_rows_cell = x.max()
                ost_rows = ost_rows - num_rows_cell - 1

                if ost_rows >= 0 and i == len(status['name']) - 1:
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'СТАТУС ВЫПОЛНЕНИЯ ПОРУЧЕНИЙ ГЕНЕРАЛЬНОГО ДИРЕКТОРА В РАМКАХ ПРОЕКТА СТРОИТЕЛЬСТВА', subinf)
                    head2(shapes)
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(2.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(10.63)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{status['name'][j]}" + "\n" * int(num_rows_cell - num_rows_name + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(0.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(2)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f'{j + 1}' + "\n" * int(num_rows_cell + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(13.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{status['complete_date'][j]}" + "\n" * int(num_rows_cell - num_rows_complete_date + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(20.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['people'][j] == None or status['people'][j] == "nan":
                            status['people'][j] = " "
                        p.text = f"{status['people'][j]}" + "\n" * int(num_rows_cell - num_rows_people + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(27.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(6)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['status'][j] == None or status['status'][j] == "nan":
                            status['status'][j] = " "
                        p.text = f"{status['status'][j]}" + "\n" * int(num_rows_cell - num_rows_status + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                        
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(33.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['comments'][j] == None or status['comments'][j] == "nan":
                            status['comments'][j] = " "
                        p.text = f"{status['comments'][j]}" + "\n" * int(num_rows_cell - num_rows_comments + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                        
                elif ost_rows < 0:
                    ost_rows = 23
                    slide = ppt.slides.add_slide(blank_slide_layout)
                    shapes = slide.shapes
                    subinf = ""
                    name_of_slide(slide, f'СТАТУС ВЫПОЛНЕНИЯ ПОРУЧЕНИЙ ГЕНЕРАЛЬНОГО ДИРЕКТОРА В РАМКАХ ПРОЕКТА СТРОИТЕЛЬСТВА', subinf)
                    head2(shapes)
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(2.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(10.63)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{status['name'][j]}" + "\n" * int(num_rows_cell - num_rows_name + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(0.41)
                    top = Cm(4.55 - 0.61)
                    width = Cm(2)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f'{j + 1}' + "\n" * int(num_rows_cell + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(13.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        p.text = f"{status['complete_date'][j]}" + "\n" * int(num_rows_cell - num_rows_complete_date + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(20.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['people'][j] == None or status['people'][j] == "nan":
                            status['people'][j] = " "
                        p.text = f"{status['people'][j]}" + "\n" * int(num_rows_cell - num_rows_people + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)

                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(27.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(6)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['status'][j] == None or status['status'][j] == "nan":
                            status['status'][j] = " "
                        p.text = f"{status['status'][j]}" + "\n" * int(num_rows_cell - num_rows_status + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                        
                    # назначаем параметры текстовой ячейки заголовка слайда
                    left = Cm(33.04)
                    top = Cm(4.55 - 0.61)
                    width = Cm(7)
                    height = Cm(22.86 - 4.55)
                        
                    # создаем текстовую ячейку
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    for j in range(ind, i + 1):
                        
                        # разрешаем перенос слов
                        tf.word_wrap = True
                        
                        # добавляем параграф текста
                        p = tf.add_paragraph()
                        if status['comments'][j] == None or status['comments'][j] == "nan":
                            status['comments'][j] = " "
                        p.text = f"{status['comments'][j]}" + "\n" * int(num_rows_cell - num_rows_comments + 1)
                        # выравниваем по центрк по вертикали
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        # выравниваем по центру по горизонтали
                        p.alignment = PP_ALIGN.LEFT
                        p.font.size = Pt(18)
                        p.font.name = "Arial Narrow"
                        p.font.color.rgb = RGBColor(32, 56, 100)
                    ind = i + 1
            
            # оформление всех слайдов
            cnt = 0
            for slide in ppt.slides:
                if cnt != 0 and cnt != app_flag:
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

                    pic = slide.shapes.add_picture("system_photo/icon2.png", Cm(0.41), Cm(0.35), width=Cm(2.38), height=Cm(2.53))
                cnt += 1
                
            # сохранение
            binary_output = BytesIO()
            ppt.save(binary_output)
            st.write('Генерация прошла успешно!')

            st.download_button(label='Скачать сгенерированную презентацию',
                               data=binary_output.getvalue(),
                               file_name='presentation.pptx')
            
            
main()