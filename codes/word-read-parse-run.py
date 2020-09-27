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

# print("gutterAtTop: ", document.settings.gutter_at_top)
# print("styles: ", (document.styles))

for para in all_paras[:1]:
    # pass
    for r0 in para.runs[:4]:
       
        print("r.text:", r0.text)
        print("r.style:", r0.style.name)
        print("r.font.asii:", r0.font.name, r0.style.font.name)
        print("r.font.asia:", r0.font.name_eastasia)
        print("r.font.size:", r0.font.size, r0.style.font.size)
        print("r.font.rgb:", r0.font.color.rgb)
        print("r.font.theme_color:", r0.font.color.theme_color)

        # if(r0.font.bold is not r0.bold): 
        print("r.bold:", r0.font.bold, r0.style.font.bold)
        # if(r0.font.italic is not r0.italic):
        print("r.italic:", r0.font.italic, r0.style.font.italic)
        # if(r0.font.underline is not r0.underline):
        print("r.underline:", r0.font.underline, r0.style.font.underline)
        print()

    print("para.text:", para.text)
    print("para.style:", para.style.name)

    ##  paragraph_format 
    # p_format = para.paragraph_format
    # pstyle_format = para.style.paragraph_format
    # # para.alignment === para.paragraph_format.alignment
    # print("para.alignment:", p_format.alignment, pstyle_format.alignment)
    # print("para.first_line_indent:", p_format.first_line_indent.cm, pstyle_format.first_line_indent)
    # print("para.left_indent:", p_format.left_indent.cm, pstyle_format.left_indent)
    # print("para.right_indent:", p_format.right_indent.cm, pstyle_format.right_indent)
    # print("para.space_before:", p_format.space_before.cm, pstyle_format.space_before)
    # print("para.space_after:", p_format.space_after.cm, pstyle_format.space_after)
    # print("para.line_spacing:", p_format.line_spacing, pstyle_format.line_spacing)
    # print("para.line_spacing_rule:", p_format.line_spacing_rule, pstyle_format.line_spacing_rule)
    # print("para.page_break_before:", p_format.page_break_before, pstyle_format.page_break_before)
    # print("para.keep_with_next:", p_format.keep_with_next, pstyle_format.keep_with_next)
    # print("para.keep_together:", p_format.keep_together, pstyle_format.keep_together)
    # print("para.widow_control:", p_format.widow_control, pstyle_format.widow_control)
    # print("para.firstchardropcap:", para.paragraph_format.first_char_dropcap)
    # print("para.firstchardropcaplines:", para.paragraph_format.first_char_dropcap_lines)

    ## font
    print("para.asii:", para.style.font.name)
    print("para.asia:", para.style.font.name_eastasia)
    print("para.size:", para.style.font.size)
    print("para.bold:", para.style.font.bold)
    print("para.italic:", para.style.font.italic)

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