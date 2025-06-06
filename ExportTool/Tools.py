import os.path  # 根据内容，获得编号及内容的字典
import re

import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# 定义命名空间
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

#段落的背景颜色
def get_paragraph_shading(paragraph):
    p = paragraph._element  # 获取段落的底层 XML 元素

    # 查找 w:pPr（段落属性）元素
    pPr = p.find(f'{WORD_NAMESPACE}pPr')

    if pPr is not None:
        # 查找 w:shd 元素
        shd = pPr.find(f'{WORD_NAMESPACE}shd')

        if shd is not None:
            # 获取填充颜色
            val = shd.get(f'{WORD_NAMESPACE}fill')
            if val == 'auto':
                return None
            return val
    return None


# from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# 复制段落
def copy_paragraph(source_para ,new_para):
    # 复制段落格式
    _copy_paragraph_format(source_para, new_para)

    # 复制所有Run及其格式
    for source_run in source_para.runs:
        new_run = new_para.add_run(source_run.text)
        _copy_run_format(source_run, new_run)


def _copy_paragraph_format(source, target):
    """深度复制段落级格式"""
    # 对齐方式
    target.paragraph_format.alignment = source.paragraph_format.alignment

    # 缩进设置
    target.paragraph_format.left_indent = source.paragraph_format.left_indent
    target.paragraph_format.right_indent = source.paragraph_format.right_indent
    target.paragraph_format.first_line_indent = source.paragraph_format.first_line_indent

    # 间距设置
    target.paragraph_format.space_before = source.paragraph_format.space_before
    target.paragraph_format.space_after = source.paragraph_format.space_after
    target.paragraph_format.line_spacing = source.paragraph_format.line_spacing

    # 分页控制
    target.paragraph_format.keep_together = source.paragraph_format.keep_together
    target.paragraph_format.keep_with_next = source.paragraph_format.keep_with_next
    target.paragraph_format.page_break_before = source.paragraph_format.page_break_before

    # 样式设置（谨慎使用）
    if source.style:
        try:
            target.style = source.style
        except KeyError:
            pass  # 目标文档无此样式时跳过


def _copy_run_format(source, target):
    """深度复制Run级格式"""
    # 字体基础
    target.font.name = 'Times New Roman'
    #target.font.size = source.font.size

    # 设置中文字体
    # 需导入 qn 模块
    from docx.oxml.ns import qn
    # run_2.font.name = '楷体'  # 注：如果想要设置中文字体，需在前面加上这一句
    target.font.element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
    # 设置字体大小
    target.font.size = Pt(14)

    # 字体样式
    target.font.bold = source.font.bold
    target.font.italic = source.font.italic
    target.font.underline = source.font.underline
    target.font.strike = source.font.strike

    # 颜色设置
    if source.font.color.rgb:
        target.font.color.rgb = source.font.color.rgb
    if source.font.highlight_color:
        target.font.highlight_color = source.font.highlight_color

    # 高级格式
    target.font.subscript = source.font.subscript
    target.font.superscript = source.font.superscript

    # 字间距
    #if source.font.kern:
    #    target.font.kern = source.font.kern

    # 字符缩放
    # if source.font.scaling:
    #     target.font.scaling = source.font.scaling


def split_options(text):
    text = text.replace("．", ".")
    arr = ["A", "B", "C", "D", "E", "F"]
    text.replace("．", ".")
    rtnArr = []
    sIdx = -1
    curDesc = ""
    isLastLatterIsEmpty: bool = True
    for lIdx in range(len(text)):
        curLeter = text[lIdx]
        if curLeter in arr and lIdx < len(text) - 1 and text[lIdx + 1] == "." and isLastLatterIsEmpty and arr.index(
                curLeter) == sIdx + 1:
            # 开始
            if curDesc != "" and len(curDesc) > 1:
                rtnArr.append(curDesc)
            curDesc = "" + curLeter
            sIdx = sIdx + 1
        else:
            if curDesc == '':
                continue
            curDesc = curDesc + curLeter
        isLastLatterIsEmpty = curLeter == " " or curLeter != "\t" or curLeter != "\n"
    if len(curDesc) > 1:
        rtnArr.append(curDesc)
    return rtnArr


# 替换 (1)为1.
def replace_parentheses_with_period(text):
    # 定义正则表达式模式
    pattern = r'\((\d+)\)'
    # 使用 re.sub 替换匹配到的内容
    replaced_text = re.sub(pattern, r'\1.', text)

    return replaced_text


def getNumPairs(fromId: int, content: str, end: str = None, max: int = -1):
    if fromId <= 0:
        print("getNumPairs ID错误，%d" % fromId)
        return None

    idx = checkStrNumFormat(fromId, content)
    if idx == -1:
        return None
    idx = idx + len(str(fromId)) + 1
    if fromId >= max != -1:
        return content[idx:]
    times = 1
    idx1 = checkStrNumFormat(fromId + 1, content, times)
    # 为了解决解析中也有数字+点，特别是 1.xx4.xx 2.dddd 3.eee 4.xx不会被误识别
    while idx1 != -1 and idx1 < idx:
        times = times + 1
        idx1 = checkStrNumFormat(fromId + 1, content, times)
    if idx1 == -1:
        rtnStr = content[idx:]
        if end and fromId + 1 == max:
            endIdx = rtnStr.find(end)
            if endIdx != -1:
                return rtnStr[0:endIdx]
        return rtnStr
    return content[idx:idx1]


def checkStrNumFormat(fromId: int, content: str, newN: int = -1):
    _n = 1
    if newN > 0:
        _n = newN
    lastIdx = -1
    while True:
        idx = solve(content, ("%d." % fromId), _n)
        if idx == -1 or (lastIdx != -1 and idx > lastIdx):
            return lastIdx
        if idx > 0:
            exNumStr = content[idx - 1:idx]
            if exNumStr.isdigit():
                _n += 1
                continue
        lastIdx = idx
        _n += 1


#字符第几次出现
def solve(s, str, n):
    sep = s.split(str, n)
    if len(sep) <= n:
        return -1
    return len(s) - len(sep[-1]) - len(str)


def getImportPath():
    path = "导入.txt"
    if not os.path.exists(path):
        raise Exception("导入文件不存在%s" % path)
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
        content = content.strip()
        return content


#获得【数字】开头的部分
def isReadingMain(content: str):
    # 定义正则表达式模式
    pattern = r'^【(\d+)】'
    # 使用 re.match 来查找匹配项
    match = re.match(pattern, content.strip())
    if match:
        #num:str = match.group(1)  # 返回捕获的第一组内容
        return True
    else:
        return False


def getContentBeforeNumAndPot(content: str):
    #一定要包含点
    if content.find(".") == -1:
        return content
    # 定义正则表达式模式
    # 匹配从字符串的开头到第一个数字加点之前的所有内容
    pattern = r'^(.*?)(?=\d+\.)'

    # 使用 re.search 查找第一个匹配项
    match = re.search(pattern, content)

    if match:
        # 返回匹配的内容
        return match.group(1).strip()
    else:
        # 如果没有找到匹配项，则返回 None 或者空字符串
        return content



#获得第一个数字+点格式中的数字
def get_first_number_before_dot(text):
    # 定义正则表达式模式
    # 匹配数字加点
    pattern = r'(\d+)\.'

    # 使用 re.search 查找第一个匹配项
    match = re.search(pattern, text)

    if match:
        # 返回匹配的数字
        rtn = int(match.group(1))
        if rtn == 0:  # 如果是0.几这种，直接去掉
            text = text.replace("0.","")
            return get_first_number_before_dot(text)
        idx:int = text.find(str(rtn)+".") + len(str(rtn)+".") - 1
        if idx + 1 < len(text):
            if text[idx+1].isdigit():
                text = text.replace( str(rtn) + "." + str(text[idx+1]) , str(rtn) + " ." + str(text[idx+1]) )
                return get_first_number_before_dot(text)
        return rtn
    else:
        # 如果没有找到匹配项，则返回 None 或者空字符串
        return None


# 从指定数字开始，最大的数值
def get_max_number( text: str) :
    from_num:int = get_first_number_before_dot(text)
    if from_num is None:
        return None
    #最多检查100
    last_num:int = from_num
    cur_num:int = from_num
    for i in range(100):
        num_str:str = str(cur_num) + "."
        idx:int = text.find( num_str )
        if idx == -1 :
            #下一位不是数字
            # if i < 100 - 1:
                #if not text[i+1].isdigit():
                    return last_num
                #else:
                 #   print("sss")
            #return last_num
        #如果有包含1.2类似的数字，则直接替换
        if text[idx + len(num_str)].isdigit() and not is_circle_num(text[idx + len(num_str)]):
            # 点先替换成空格
            str_list = list(text)
            str_list.insert(idx + (len)(num_str) - 1,' ')
            text = "".join(str_list)
            #text = text.replace(num_str,str(cur_num) + " .")
        else:
            last_num = cur_num
            cur_num = cur_num + 1

def is_circle_num(s:str):
    a = ["①","②","③"]
    for a_s in a :
        if s.find(a_s) != -1:
            return True
    return False

def is_start_with_num_and_point(s:str):
    s = s.strip()
    if s.__contains__("."):
        arr = s.split(".")
        if arr[0].isdigit():
            return True
        return False
    else:
        return False


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def read_table(table):
    arr = [[cell.text for cell in row.cells] for row in table.rows]
    content = []
    for a in arr:
        for a1 in a :
            content.append(a1)
            #content = content + a1
    c = ",".join(content)
    return c

#将指定内容，按underlines添加runs
def add_underlines_runs_by_arr(para:Paragraph,content:str):
    final_arr:[] = []
    tmp_part:str = ""
    wait_flag:int = 0 # 0- 正常  1 - 等待匹配
    has_underline:bool = False
    need_write:bool = False
    content_len:int = len(content)
    for i in range(content_len):
        letter = content[i]
        if letter == "^" and wait_flag == 0 :
            wait_flag = 1
            has_underline = False
            need_write = True
        elif letter == "^" and wait_flag == 1 :
            wait_flag = 0
            has_underline = True
            need_write = True
        if letter != "^" :
            tmp_part = tmp_part + letter
        if content_len - 1 == i :
            need_write = True
        if need_write:
            r = para.add_run(tmp_part)
            r.font.name = 'Times New Roman'
            r.font.name = 'Times New Roman'  # 注：这个好像设置 run 中的西文字体
            # 设置中文字体
            # 需导入 qn 模块
            from docx.oxml.ns import qn
            # run_2.font.name = '楷体'  # 注：如果想要设置中文字体，需在前面加上这一句
            r.font.element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
            # 设置字体大小
            r.font.size = Pt(14)
            r.font.underline = has_underline
            tmp_part = ""
            need_write = False
            has_underline = False




if __name__ == '__main__':
    test: str = ('''51．考查名词。句意：植物性牛奶多年来越来越受欢迎，但不同牛奶和牛奶替代品的营养成分不同，品牌之间也存在分歧。分析可知，空前是介词，所以空处应填名词，根据后文“The global dairy alternatives market is 　2　to grow from $22.25 billion in 2021 to $53.97 billion in 2028”可知，牛奶市场越来越大，所以空处应选popularity意为“受欢迎”，符合句意。故选H项。r"
                 52．考查动词。句意：根据《财富》商业观察的一份报告，全球乳制品替代品市场预计将从2021年的222.5亿美元增长到2028年的539.7亿美元。根据后文中“to $53.97 billion in 2028”，可推测，现在还未到2028年，所以应该是表预测，所以应填project意为“预测”，又本句主语是market与project之间是被动关系，应用被动语态，空前已有be动词is，空处用project的过去分词。故选F项
53．考查形容词。句意：2022年2月，世界上第一款土豆牛奶在英国推出，它标榜自己是“市场上最可持续的植物性乳制品替代品”。根据空前的the most可知空处应填形容词，表示“最......的”，此处分析选项，可知应在A，C，E，I中选，带入句子中，可知sustainable意为“可持续的”符合句意。故选C项。''')
    test = test.replace("．",".")
    opts = get_max_number(test)
    print(opts)
