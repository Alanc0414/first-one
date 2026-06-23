# -*- coding: utf-8 -*-
"""
生成问卷星最新版 Excel 导入格式
每题占一个单元格，用 <br/> 换行，整题含选项、答案、解析
"""

import re
from openpyxl import Workbook

from generate_quiz import day1_data, day2_data, day3_data

BR = '<br/>'


def clean_text(text):
    if not text:
        return ''
    text = str(text)
    text = text.replace('"', '').replace('"', '').replace('"', '')
    text = text.replace('（多选）', '').replace('(多选)', '')
    text = text.replace('______', '____')
    return text.strip()


def format_options(options_str):
    lines = []
    for line in options_str.strip().split('\n'):
        m = re.match(r'^([A-Z])\.\s*(.+)', line.strip())
        if m:
            lines.append(f'{m.group(1)}. {clean_text(m.group(2))}')
    return BR.join(lines)


def format_question_stem(question, q_type):
    question = clean_text(question)
    if q_type == '多选题':
        question = re.sub(r'[？?]\s*$', '', question)
        if not question.endswith('（ ）'):
            question = question + '（ ）'
    return question


def format_fill_answer(answer):
    return answer.replace(';', '；')


def build_question_cell(num, row):
    _, q_type, question, options_str, answer, source, explanation = row
    question = format_question_stem(question, q_type)
    parts = [f'{num}. {question}']

    if q_type in ('单选题', '多选题'):
        parts.extend(['', format_options(options_str), '', f'答案：{answer}'])
    else:
        analysis = clean_text(f'依据：{source}。{explanation}')
        parts.extend([
            '',
            f'答案：{format_fill_answer(answer)}',
            '',
            f'解析：{analysis}',
            '',
        ])

    return BR.join(parts)


def create_wjx_excel(filename, data):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    ws.cell(row=1, column=1, value='题干')

    for idx, row in enumerate(data, 1):
        ws.cell(row=idx + 1, column=1, value=build_question_cell(idx, row))

    ws.column_dimensions['A'].width = 100
    wb.save(filename)
    print('已生成：' + filename)


if __name__ == '__main__':
    create_wjx_excel('CMport_问卷星导入_Day1.xlsx', day1_data)
    create_wjx_excel('CMport_问卷星导入_Day2.xlsx', day2_data)
    create_wjx_excel('CMport_问卷星导入_Day3.xlsx', day3_data)
    print('\n已按问卷星最新版单单元格格式生成（共3个，各10题）。')
