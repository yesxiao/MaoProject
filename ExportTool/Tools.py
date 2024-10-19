import os.path  # 根据内容，获得编号及内容的字典
import re


def split_options(text):
    text = text.replace("．",".")
    arr = ["A","B","C","D","E","F"]
    text.replace("．",".")
    rtnArr = []
    sIdx = -1
    curDesc = ""
    isLastLatterIsEmpty:bool = True
    for lIdx in range(len(text)):
        curLeter = text[lIdx]
        if curLeter in arr and lIdx < len(text) - 1 and text[lIdx+1] == "." and isLastLatterIsEmpty and arr.index(curLeter) == sIdx + 1:
            # 开始
            if curDesc != "" and len(curDesc) > 1 :
                rtnArr.append(curDesc)
            curDesc = "" + curLeter
            sIdx = sIdx + 1
        else:
            curDesc = curDesc + curLeter
        isLastLatterIsEmpty = curLeter == " " or curLeter != "\t" or curLeter != "\n"
    if len(curDesc) > 1 :
        rtnArr.append(curDesc)
    return rtnArr



# 替换 (1)为1.
def replace_parentheses_with_period(text):
    # 定义正则表达式模式
    pattern = r'\((\d+)\)'
    # 使用 re.sub 替换匹配到的内容
    replaced_text = re.sub(pattern, r'\1.', text)

    return replaced_text

def getNumPairs(fromId:int,content:str,end:str ,max:int ):
    if fromId <= 0 :
        print("getNumPairs ID错误，%d" % fromId )
        return None

    idx = checkStrNumFormat( fromId , content )
    if idx == -1:
        return None
    idx = idx + len(str(fromId)) + 1
    if fromId >= max :
        return content[idx:]
    times = 1
    idx1 = checkStrNumFormat( fromId + 1 , content , times)
    # 为了解决解析中也有数字+点，特别是 1.xx4.xx 2.dddd 3.eee 4.xx不会被误识别
    while idx1 != -1 and idx1 < idx :
        times = times + 1
        idx1 = checkStrNumFormat( fromId + 1 , content , times)
    if idx1 == -1:
        rtnStr = content[idx:]
        if end and fromId + 1 == max :
            endIdx = rtnStr.find(end)
            if endIdx != -1:
                return rtnStr[0:endIdx]
        return rtnStr
    return content[idx:idx1]

def checkStrNumFormat(fromId:int,content:str,newN:int = -1):
    _n = 1
    if newN > 0 :
        _n = newN
    lastIdx = -1
    while True :
        idx = solve(content,("%d." % fromId),_n)
        if idx == -1 or ( lastIdx != -1 and  idx > lastIdx ):
            return lastIdx
        if idx > 0 :
            exNumStr = content[idx-1:idx]
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
    path="导入.txt"
    if not os.path.exists(path):
        raise Exception("导入文件不存在%s" % path)
    with open(path,"r",encoding="utf-8") as f:
        content = f.read()
        content = content.strip()
        return content

#获得【数字】开头的部分
def isReadingMain(content:str):
    # 定义正则表达式模式
    pattern = r'^【(\d+)】'
    # 使用 re.match 来查找匹配项
    match = re.match(pattern, content.strip())
    if match:
        #num:str = match.group(1)  # 返回捕获的第一组内容
        return True
    else:
        return False




if __name__ == '__main__':
    test: str = (r"A．Taichung,helGE.lo. PAWS	B．DetrA.oi Ct Pub.lic Radio"
                 r"．UNICEF.	D．ORBIS.")
    opts = split_options(test)
    for k in opts:
        print(k )


