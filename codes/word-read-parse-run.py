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

print("gutterAtTop: ", document.settings.gutter_at_top)

for para in all_paras[:1]:
    # pass
    for r0 in para.runs[:4]:
       
        # print("r.text:", r0.text)
        # print("r.font.asii:", r0.font.name)
        # print("r.font.asia:", r0.font.name_eastasia)
        # print("r.font.size:", r0.font.size)
        # print("r.font.rgb:", r0.font.color.rgb)
        # print("r.font.theme_color:", r0.font.color.theme_color)

        if(r0.font.bold is not r0.bold): 
            print("r.bold:", r0.font.bold, r0.bold)
        if(r0.font.italic is not r0.italic):
            print("r.italic:", r0.font.italic, r0.italic)
        # if(r0.font.underline is not r0.underline):
        print("r.underline:", r0.font.underline, r0.underline)

    # print("para.text:", para.text)
    # print("para.first_char_dropcap:", para.paragraph_format.first_char_dropcap)
    # print("para.first_char_dropcap_lines:", para.paragraph_format.first_char_dropcap_lines)
    # print("para.style:", para.style.name)
    # print("para.asii:", para.style.font.name)
    # print("para.asia:", para.style.font.name_eastasia)
    # print("para.size:", para.style.font.size)
    # print("para.bold:", para.style.font.bold)
    # print("para.italic:", para.style.font.italic)

    # for run in para.runs:
    #     print(run.text)
    #     print(run.font.name)
    #     print("-------")

t4 = time.perf_counter()

# styles =  document.styles
# for i, style in enumerate(styles):
#     pass
#     # print(i, style)
#     # print('__________')
# t5 = time.perf_counter()

print('timing:', t2-t1, t3-t2, t4-t3)