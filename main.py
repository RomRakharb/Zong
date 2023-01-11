from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4, landscape
import pandas
import os
import calendar
import time
import pathlib
import shutil


def clear():
    for f in os.listdir('tmp'):
        if not f.endswith(".pdf"):
            continue
        os.remove(os.path.join('tmp', f))


def reset():
    df = pandas.read_excel('zong.xlsx')
    name_list = []
    for i in range(len(df)):
        name_list.append(str(df.loc[i][0]))
    return name_list


def saveItem(index, data_list):
    df = pandas.read_excel('zong.xlsx')
    df.loc[index] = data_list
    df.to_excel('zong.xlsx', index=False, header=['topic', 'line1', 'line2', 'line3', 'line4', 'line5'])
    sort_df()


def deleteAction(index):
    df = pandas.read_excel('zong.xlsx')
    df.drop(index, axis=0, inplace=True)
    df.to_excel('zong.xlsx', index=False, header=['topic', 'line1', 'line2', 'line3', 'line4', 'line5'])


def selectedItem(name):
    name_list = reset()
    df = pandas.read_excel('zong.xlsx')
    receiver = df.loc[name_list.index(name)]
    return receiver


def printEnvelope(des, notSend, name):
    name_list = reset()
    df = pandas.read_excel('zong.xlsx')
    sender = df.loc[0]
    receiver = df.loc[name_list.index(des)]
    pdfmetrics.registerFont(TTFont('TH Sarabun New', 'THSarabunNew.ttf'))
    current_GMT = time.gmtime()
    ts = calendar.timegm(current_GMT)
    filename = str(name_list.index(des)) + '-' + str(ts) + '.pdf'
    c = canvas.Canvas(filename)
    c.setFont('TH Sarabun New', 16)
    c.setPageSize((683.1496063, 297.63779528))
    ySender = 200
    yReceiver = 120
    xSender = 70
    xPreReceiver = 260
    xReceiver = 310
    xStamp = 575
    yStamp = 240
    xRect = 500
    yRect = 175
    xNotsend = 500
    yNotsend = 50
    for i in sender[1:]:
        if not pandas.isna(i):
            c.drawString(xSender, ySender, i)
            ySender -= 20
    c.drawString(xPreReceiver, yReceiver, 'เรียน')
    for i in receiver[1:]:
        if not pandas.isna(i):
            c.drawString(xReceiver, yReceiver, i)
            yReceiver -= 20
    if notSend is False:
        c.rect(xRect, yRect, 150, 100)
        with open('zong.txt', encoding='utf8') as f:
            lines = f.readlines()
            for line in lines:
                c.drawCentredString(xStamp, yStamp, line.replace('\n', ''))
                yStamp -= 20
    else:
        c.drawString(xNotsend, yNotsend, '( ' + name + ' )')
    c.save()
    shutil.move(str(pathlib.Path().resolve()) + '\\' + filename,
                str(pathlib.Path().resolve()) + "\\tmp\\" + filename)
    os.startfile(str(pathlib.Path().resolve()) + "\\tmp\\" + filename)


def printA4(des, notSend, name):
    name_list = reset()
    df = pandas.read_excel('zong.xlsx')
    sender = df.loc[0]
    receiver = df.loc[name_list.index(des)]
    pdfmetrics.registerFont(TTFont('TH Sarabun New', 'THSarabunNew.ttf'))
    current_GMT = time.gmtime()
    ts = calendar.timegm(current_GMT)
    filename = str(name_list.index(des)) + '-' + str(ts) + '.pdf'
    c = canvas.Canvas(filename)
    c.setFont('TH Sarabun New', 20)
    c.setPageSize(landscape(A4))
    x_max = 842
    y_max = 595
    xSender = 50
    ySender = y_max - 150
    xPreReceiver = x_max / 2 - 100
    xReceiver = x_max / 2 - 50
    yReceiver = y_max / 2
    xStamp = x_max - 150
    yStamp = y_max - 80
    xRect = x_max - 250
    yRect = y_max - 150
    xNotsend = x_max / 2
    yNotsend = y_max / 4
    spacing = 25
    c.rect(20, 20, 802, 555)
    c.drawImage('garuda.png', 50, y_max - 175, 50, mask='auto', preserveAspectRatio=True)
    for i in sender[1:]:
        if not pandas.isna(i):
            c.drawString(xSender, ySender, i)
            ySender -= spacing
    c.drawString(xPreReceiver, yReceiver, 'เรียน')

    for i in receiver[1:]:
        if not pandas.isna(i):
            c.drawString(xReceiver, yReceiver, i)
            yReceiver -= spacing
    if notSend is False:
        c.rect(xRect, yRect, 200, 100)
        with open('zong.txt', encoding='utf8') as f:
            lines = f.readlines()
            for line in lines:
                c.drawCentredString(xStamp, yStamp, line.replace('\n', ''))
                yStamp -= spacing
    else:
        c.drawCentredString(xNotsend, yNotsend, '( ' + name + ' )')
    c.save()
    shutil.move(str(pathlib.Path().resolve()) + '\\' + filename,
                str(pathlib.Path().resolve()) + "\\tmp\\" + filename)
    os.startfile(str(pathlib.Path().resolve()) + "\\tmp\\" + filename)


def sort_df():
    df = pandas.read_excel('zong.xlsx')
    top = df.loc[0:1]
    add = df.loc[1]
    df.drop(1, axis=0, inplace=True)
    df.drop(0, axis=0, inplace=True)
    sorted_df = df.sort_values('topic')
    new_df = pandas.concat([top, sorted_df])
    new_df.to_excel('zong.xlsx', index=False, header=['topic', 'line1', 'line2', 'line3', 'line4', 'line5'])


if __name__ == "__main__":
    pass
