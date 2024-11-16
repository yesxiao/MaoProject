import os.path  # 根据内容，获得编号及内容的字典
import re


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
        if text.find( str(cur_num) + "." ) == -1:
            return last_num
        last_num = cur_num
        cur_num = cur_num + 1

def is_start_with_num_and_point(s:str):
    s = s.strip()
    if s.__contains__("."):
        arr = s.split(".")
        if arr[0].isdigit():
            return True
        return False
    else:
        return False


if __name__ == '__main__':
    test: str = (r"【导语】本文是一篇说明文。文章说明了在语言学习的中、高级阶段的单词学习法。"
                 r"31.考查动词词义辨析。句意：然而，功能性33.语言的熟练，需要掌握相当多")
    opts = get_max_number(test)
    print(opts)
