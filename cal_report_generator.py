from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import random
import datetime
import matplotlib.pyplot as plt

#Import template document
template = DocxTemplate('cal_report_template.docx')

# #Generate list of random values
# table_contents = []
# x = []
# y = []
#
# for i in range(0,12):
#     number = round(random.random(),3)
#
#     table_contents.append({
#         'Index': i,
#         'Value': number
#     })
#
#     x.append(i)
#     y.append(number)
#
# #Plot random values and save figure
# fig = plt.figure()
# plt.plot(x, y)
# fig.savefig('image.png', dpi=fig.dpi)
#
# #Import saved figure
# image = InlineImage(template,'image.png',Cm(10))


'''Declare template variables'''
context = {
    't1': '26.7',
    't1_1': '26.5',
    't1_2': '26.5',
    't1_3': '26.5',
    't1_4': '26.5',
    't1_5': '26.5',
    't1_6': '26.5',
    't1_7': '26.5',
    't1_8': '26.5',
    't1_9': '26.9',
    't1_av': 26.7,
    't1_uni': 0.1,
    't1_unc': '0.08',

    't2': '26.7',
    't2_1': '26.5',
    't2_2': '26.5',
    't2_3': '26.5',
    't2_4': '26.5',
    't2_5': '26.5',
    't2_6': '26.5',
    't2_7': '26.5',
    't2_8': '26.5',
    't2_9': '26.9',
    't2_av': 26.7,
    't2_uni': 0.1,
    't2_unc': '0.08',

    't3': '26.7',
    't3_1': '26.5',
    't3_2': '26.5',
    't3_3': '26.5',
    't3_4': '26.5',
    't3_5': '26.5',
    't3_6': '26.5',
    't3_7': '26.5',
    't3_8': '26.5',
    't3_9': '26.9',
    't3_av': 26.7,
    't3_uni': 0.1,
    't3_unc': '0.08',

    't4': '26.7',
    't4_1': '26.5',
    't4_2': '26.5',
    't4_3': '26.5',
    't4_4': '26.5',
    't4_5': '26.5',
    't4_6': '26.5',
    't4_7': '26.5',
    't4_8': '26.5',
    't4_9': '26.9',
    't4_av': 26.7,
    't4_uni': 0.1,
    't4_unc': '0.08',

    't5': '26.7',
    't5_1': '26.5',
    't5_2': '26.5',
    't5_3': '26.5',
    't5_4': '26.5',
    't5_5': '26.5',
    't5_6': '26.5',
    't5_7': '26.5',
    't5_8': '26.5',
    't5_9': '26.9',
    't5_av': 26.7,
    't5_uni': 0.1,
    't5_unc': '0.08',
}

#Render automated report
template.render(context)
template.save('generated_report.docx')