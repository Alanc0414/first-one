# -*- coding: utf-8 -*-
"""
将题库数据转换为问卷星可直接导入的 Excel 格式
列结构：题型 | 题干 | 选项A~F | 正确答案 | 分值 | 答案解析
"""

import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from generate_quiz import day1_data, day2_data, day3_data

WJX_COLUMNS = [
    '题型', '题干',
    '选项A', '选项B', '选项C', '选项D', '选项E', '选项F',
    '正确答案', '分值', '答案解析',
]

TYPE_MAP = {
    '单选题': '单选题',
    '多选题': '多选题',
    '填空题': '填空题',
}


def parse_options(options_str):
    opts = []
    for line in options_str.strip().split('\n'):
        m = re.match(r'^[A-Z]\.\s*(.+)', line.strip())
        if m:
            opts.append(m.group(1))
    return opts


def convert_fill_question(question, answer_str):
    answers = [a.strip() for a in answer_str.replace('；', ';').split(';')]
    result = question
    for ans in answers:
        result = result.replace('______', '{' + ans + '}', 1)
    return result, answers


def is_multi_blank(question):
    return question.count('______') > 1


def convert_row(row, score=10):
  # row: [题号, 题型, 题目, 选项, 正确答案, 来源, 解析]
    _, q_type, question, options_str, answer, source, explanation = row

    opts = parse_options(options_str) if options_str != '（填空）' else []
    opt_cols = (opts + [''] * 6)[:6]

    if q_type == '填空题':
        question, answers = convert_fill_question(question, answer)
        wjx_type = '多项填空' if is_multi_blank(row[2]) else '填空题'
        wjx_answer = ';'.join(answers)
    else:
        wjx_type = TYPE_MAP[q_type]
        wjx_answer = answer

    analysis = f'【依据】{source}\n{explanation}'

    return [wjx_type, question] + opt_cols + [wjx_answer, score, analysis]


def create_wjx_excel(filename, data):
    wb = Workbook()
    ws = wb.active
    ws.title = '试题'

    header_font = Font(name='微软雅黑', bold=True, size=11)
    header_fill = PatternFill('solid', fgColor='D9E1F2')
    border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF'),
    )

    for col_idx, col_name in enumerate(WJX_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    for row_idx, row in enumerate(data, 2):
        converted = convert_row(row)
        for col_idx, val in enumerate(converted, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name='微软雅黑', size=10)
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=True)

    widths = [10, 42, 22, 22, 22, 22, 22, 22, 12, 6, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'A2'
    wb.save(filename)
    print('已生成：' + filename)


if __name__ == '__main__':
    create_wjx_excel('CMport_问卷星导入_Day1.xlsx', day1_data)
    create_wjx_excel('CMport_问卷星导入_Day2.xlsx', day2_data)
    create_wjx_excel('CMport_问卷星导入_Day3.xlsx', day3_data)
    print('\n问卷星导入文件全部生成完毕（共3个）。')
