
 # 根据内容，获得编号及内容的字典
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
        if end :
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


if __name__ == '__main__':
    test: str = r"【分析】47.读图可知，珠穆朗玛峰的海拔为8848.86米，章子峰海拔为7543米，故二者高差是1305.86米，故A错；从前进营地前往突击营地先由东北向西南，再由西北向东南，故B错；由于珠穆朗玛峰海拔高，气温低，空气稀薄，故面临着严寒、缺氧的困难，故C正确；由图可知，该图的等高距是400米，突击营地的海拔约8400米，前进营地的海拔约6600米，故D错，故依题意选C。\
48.现代科学研究表明，海陆是不断变迁的，喜马拉雅山上发现海洋生物化石说明这里曾经是海洋，由于地壳的变动，海洋演变为陆地，故B正确，故选B。\
49.由题目可知，2018年底，相关部门发出公告称，禁止任何单位和个人进入珠穆朗玛峰国家级自然保护区绒布寺核心区域旅游，这一措施的主要目的是更好保护当地的生态环境，因为这里海拔高，气候恶劣，生态环境脆弱，故B正确，故选B。\
【点睛】中国登山队两次登上世界最高峰的壮举，具有非凡的历史意义。珠峰高程测量的核心目标是精确测定珠峰高度，测量成果可用于地球动力学板块运动等领域研究。精确的峰顶雪深、气象、风速等数据，将为冰川监测、生态环境保护等方面的研究提供第一手资料。【aaa】"
    p = getNumPairs(47, test,"【")
    print( "47:" + p)
    p = getNumPairs(48, test,"【")
    print("48:" + p)
    p = getNumPairs(49, test,"【")
    print("49:" + p )


