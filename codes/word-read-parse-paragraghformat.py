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

for para in all_paras[:3]:
    # pass
    print("para.text:", para.text)
    print("para.alignment:", para.paragraph_format.alignment)
    print("para.indent:", para.paragraph_format.left_indent, 
    para.paragraph_format.right_indent)
    print("para.space:", para.paragraph_format.space_before, 
    para.paragraph_format.space_after)
    print("para.line_spacing:", para.paragraph_format.line_spacing_rule, 
    para.paragraph_format.line_spacing)
    print("para.first_line_indent:", para.paragraph_format.first_line_indent)
    print("para.page:", para.paragraph_format.keep_together,
    para.paragraph_format.keep_with_next,
    para.paragraph_format.page_break_before,
    para.paragraph_format.widow_control)


t4 = time.perf_counter()

# styles =  document.styles
# for i, style in enumerate(styles):
#     pass
#     # print(i, style)
#     # print('__________')
# t5 = time.perf_counter()

print('timing:', t2-t1, t3-t2, t4-t3)