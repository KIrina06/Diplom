### импорт библиотек
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import datetime
from datetime import date
import pandas as pd

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
from func import know_size, check_width, check_height, make_width, make_height

### функции
# подгон текста под размеры
def resize_text(rect_height, rect_width, text, fontsize):
    
    #text = "ИНВЕНТАРИЗАЦИЯ ПРОЕКТА СООРУЖЕНИЕ АЭС «РУППУР» НА ПРЕДМЕТ НАЛИЧИЯ РИСКОВ СРЫВА ПОnnnnnnnnnnnnnnnnСТАВОК ОБОРУДОВАНИЯ И КОМПЛЕКТУЮЩИХ ИЗ 3-Х СТРАН"
    fig = plt.figure(figsize = (rect_width, rect_height))
    
    gs = GridSpec(ncols = 1, nrows = 1, figure = fig, wspace=0, hspace=0)
    ax = plt.subplot(gs[: , :])
    # изменяем текст (переносим строки или уменьшаем шрифт) так, чтобы он влезал в график
    text, fontsize = make_width(fig, ax, text, fontsize)
    text, fontsize = make_height(fig, ax, text, fontsize)
   
    return text, fontsize

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

# отрисовка таблицы
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

# отрисовка шапки таблицы рисков
def head(shapes):
    rows = 1
    cols = 4
    left = Cm(0.41)
    top = Cm(3.05)
    width = Cm(39.63)
    height = Cm(1.5)
    table = shapes.add_table(rows, cols, left, top, width, height).table
        
    cell = table.cell(0, 0)
        
    cell.text = "№"
        
    cell = table.cell(0, 1)
        
    cell.text = "НАИМЕНОВАНИЕ РИСКА"
    
    cell = table.cell(0, 2)
        
    cell.text = "ПРЕДЛАГАЕМЫЕ КОМПЕНСИРУЮЩИЕ МЕРОПРИЯТИЯ ИЛИ ПОТРЕБНОСТЬ В ПРИНЯТИИ РЕШЕНИЯ РУКОВОДСТВОМ ГК «РОСАТОМ»"
        
    cell = table.cell(0, 3)
        
    cell.text = "СТАТУС ИСПОЛНЕНИЕ КОМПЕНСИРУЮЩЕГО МЕРОПРИЯТИЯ"
        
    
    cnt = 0
    # set column widths
    for col in table.columns:
        if cnt == 0:
            col.width = Cm(2)
        elif cnt == 1:
            col.width = Cm(13.63)
        else:
            col.width = Cm(12)
        cnt += 1
        
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(12)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(218, 232, 242)

# отрисовка шапки таблицы статуса выполнения поручений
def head2(shapes):
    rows = 1
    cols = 6
    left = Cm(0.41)
    top = Cm(3.05)
    width = Cm(39.63)
    height = Cm(1.5)
    table = shapes.add_table(rows, cols, left, top, width, height).table
        
    cell = table.cell(0, 0)
        
    cell.text = "№"
        
    cell = table.cell(0, 1)
        
    cell.text = "ПОРУЧЕНИЕ"
    
    cell = table.cell(0, 2)
        
    cell.text = "СРОК ВЫПОЛНЕНИЯ"
        
    cell = table.cell(0, 3)
        
    cell.text = "ОТВ."

    cell = table.cell(0, 4)
        
    cell.text = "СТАТУС"

    cell = table.cell(0, 5)
        
    cell.text = "КОММЕНТАРИИ"
    
    cnt = 0
    # set column widths
    for col in table.columns:
        if cnt == 0:
            col.width = Cm(2)
        elif cnt == 1:
            col.width = Cm(10.63)
        elif cnt == 4:
            col.width = Cm(6)
        else:
            col.width = Cm(7)
        cnt += 1
        
    for row in table.rows:
        for cell in row.cells:
            cell.text_frame.paragraphs[0].vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.paragraphs[0].font.size = Pt(12)
            cell.text_frame.paragraphs[0].alignment = pptx.enum.text.PP_ALIGN.CENTER
            cell.text_frame.paragraphs[0].font.name = "Arial Narrow"
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(89, 89, 89)
                
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(218, 232, 242)