import re

import MyParagraph
from Tools import split_options


#解析文档
def parse_doc(doc:MyParagraph.DocAnalyse,start_str:[str]):
    last_color = "MainColor"
    datas = []
    data = None
    state:str = None
    is_valid:bool = False
    que_num:int = 0
    # main , qus , option
    for p in doc.paragraph_arr:
        if not is_valid :
            if p.paragraph.text.startswith("//题目开始"):
                is_valid = True
                continue
        if not is_valid :
            continue
        if p.paragraph.text.startswith("//题目结束"):
            break
        start_num: int = check_start_with_num(p)
        in_color = p.color and last_color
        if not p.color and last_color:
            state = "main"
            data = {}
            datas.append(data)
            #主题 直接以数字开头，说明
            if start_num > 0:
                #先加到主题里
                data[state] = []
                data[state].append(p)
                state = "qus" + str(start_num)
                que_num = start_num
        else:
            if in_color or start_num == 0 : #不是数字开头
                #选项
                is_valid_state:bool = False
                for s in start_str:
                    if check_start_with_str(p,s):
                        state = s
                        is_valid_state = True
                        break
                if not in_color and not is_valid_state and ("qus" + str(que_num) in data) and len(split_options(p.paragraph.text)) > 0 :
                    state = "option" + str(que_num)
            else:
                #问题
                state = "qus" + str(start_num)
                que_num = start_num
        if state not in data :
            data[state] = []
        data[state].append(p)
        last_color = p.color
    return datas

#判断段落是否以start_with字符串开头
def check_start_with_str(p: MyParagraph.ParagraphsAnlyse, start_with: str):
    if len(p.contents) > 0 :
        content: str = p.contents[0].content
        content = content.strip()
        if content.startswith(start_with):
            return True
    return False

#判断是否以数字开头
def check_start_with_num(p:MyParagraph.ParagraphsAnlyse ):
    if len(p.contents) > 0 :
        content:str = p.contents[0].content
        start_with_num:int = starts_with_positive_digit_and_dot(content)
        if start_with_num > 0 :
            return start_with_num
        return 0
    return 0

#是否以数字+点开头 ,如果是则返回数字，如果不是，则返回0
def starts_with_positive_digit_and_dot(s):
    """
    检查字符串是否以一个或多个数字（数值大于0）+.开头。

    参数:
    s (str): 要检查的字符串

    返回:
    bool: 如果字符串以一个或多个数字（数值大于0）+.开头返回 True，否则返回 False
    """
    # 正则表达式 ^ 表示字符串的开始
    # \d+ 匹配一个或多个数字
    # \. 匹配一个点号（需要转义）
    pattern = r'^(\d+)\.'

    match = re.match(pattern, s)
    if match:
        # 获取匹配到的第一个捕获组，即数字部分
        num_str = match.group(1)
        # 将字符串转换为整数并检查是否大于0
        n:int = int(num_str)
        return n
    else:
        return 0


if __name__ == '__main__':
    doc_path = r"D:\猫猫\0905\肖鲲测试\【改5】【答案】【选择题】中国地理：河西走廊（1）1~100（题完）_河西走廊（1）1~100（题完）_理解题.docx"
    doc_analyse = MyParagraph.DocAnalyse(doc_path)
    result_of_parse_doc = parse_doc(doc_analyse,["【答案】","【知识点】","【解析】","【详解】","【点睛】"])
    print(result_of_parse_doc)