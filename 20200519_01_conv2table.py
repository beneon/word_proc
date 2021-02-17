"""
从word版本的选择题

1. esophagus
A. 食管	B. 十二指肠	C. 胃
D. 气管	E. 胰腺
ans: A

转换成下面的形式

试题类型	单选题
题目	esophagus
分值	1
难易度	难
正确答案	A
A	选择1
B	选择2
C	选择3
D	选择4
E   选择5
答案解析	解析内容
"""

import re
import os

txt_src_file = os.path.join('txt_src','doc4.txt')
with open(txt_src_file,encoding='utf8',mode='r') as txt_src:
    txt = txt_src.read()

# 首先使用空格拆分成10份
reEmptyLineBreak = re.compile(r'\n\s*\n')
#1. esophagus\nA. 食管\tB. 十二指肠\tC. 胃\nD. 气管\tE. 胰腺\nans: A
reComponent = re.compile(r"""
    \d+\.\s+(.*)[\t\n]
    [A-Z]\.\s?(.*)[\t\n]
    [A-Z]\.\s?(.*)[\t\n]
    [A-Z]\.\s?(.*)[\t\n]
    [A-Z]\.\s?(.*)[\t\n]
    [A-Z]\.\s?(.*)[\t\n]
    ans:\s([A-Z])
    """,re.VERBOSE)
questions = reEmptyLineBreak.split(txt)

assert len(questions)==25, f'split fail: {questions}'
# 对每一份提取题干，选项和答案数据
def question_data_extract(question_str:str):
    """
    :param question_str:题目
    :return: 数据(dict) or exception
    """
    mo = reComponent.match(question_str)
    if mo:
        rst = {
            'title' : mo.group(1),
            'opt1' : mo.group(2),
            'opt2' : mo.group(3),
            'opt3' : mo.group(4),
            'opt4' : mo.group(5),
            'opt5' : mo.group(6),
            'ans' : mo.group(7),
        }
        return rst
    else:
        raise Exception(f'no match for {question_str} in question_data_extract')

question_components = [question_data_extract(e) for e in questions]

# 将dict转换成题目的形式，输出成txt文档宫复制粘贴

def component_combine(components:dict):
    rst_txt = f"""试题类型\t单选题
题目\t{components['title']}
分值\t2
难易度\t难
正确答案\t{components['ans']}
A\t{components['opt1']}
B\t{components['opt2']}
C\t{components['opt3']}
D\t{components['opt4']}
E\t{components['opt5']}
答案解析\t解析内容
"""
    return rst_txt

questions_conv = [component_combine(e) for e in question_components]

output_txt = "\n".join(questions_conv)

with open(os.path.join('txt_src','doc4_out.txt'),mode='w',encoding='utf8') as file_write:
    file_write.write(output_txt)
