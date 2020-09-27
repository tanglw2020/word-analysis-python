# -*- coding:utf-8 -*-

from docx import Document
from docx.shared import Inches
import time

t1 = time.perf_counter()
document = Document('./黄庭坚.docx')
t2 = time.perf_counter()

all_paras  = document.paragraphs
print('all_paras', len(all_paras))
t3 = time.perf_counter()

sections = document.sections
print('sections', len(sections))

section= sections[-1]
print('section.start_type:', section.start_type)
print('section.orientation:', section.orientation)
print('section.page_height:', section.page_height)
print('section.page_width:', section.page_width)
print('section.margin:', section.top_margin,
section.bottom_margin,
section.left_margin,
section.right_margin)
print("gutterAtTop: ", document.settings.gutter_at_top)
print('section.gutter:', section.gutter)
print('section.distance:', section.header_distance, section.footer_distance)


t4 = time.perf_counter()


print('timing:', t2-t1, t3-t2, t4-t3)