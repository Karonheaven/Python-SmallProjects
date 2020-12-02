# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    infile = "3-分割作者&关键词.xlsx"
    outfile = "5-作者加权统计.xlsx"
    
    # 获取源Excel路径(filepath)
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    # 利用pandas打开Excel并读取数据，所得数据写入dataframe中
    dataframe = pandas.read_excel(infile_path, sheet_name=0, header=0)
    
    # +---------------Begin---------------+
    
    # 统计dataframe中一共有多少行(nrows)，多少列(ncolumns)
    nrows = dataframe.shape[0]
    ncolumns = dataframe.shape[1]
    
    # 没有现成的函数，就转换成二维数组，直接用基础语句解决
    array_2d = dataframe.values
    
    # 参照VBA处理模式进行处理
    # 创建一个空字典
    dAuthorDict = {}
    # 从第1行开始处理，每行处理一次
    for i in range(0, nrows):
        # 统计这篇文献一共有多少作者
        iAuthorNumber = 0
        # 权重总和(分母)
        iSum = 0
        for j in range(0, ncolumns):
            # 取单元格内的作者
            sAuthorTemp = array_2d[i][j]
            if str(sAuthorTemp) == "nan":
                # 如果遇到空值，则说明作者名单已经到头
                iAuthorNumber = j
                iSum = int(j * (j + 1) / 2)
                break
        
        # 从第一个作者开始遍历
        for j in range(0, iAuthorNumber):
            # 取单元格内的作者
            sAuthorTemp = array_2d[i][j]
            # 计算单个作者的权重
            fAuthorWeight = (iAuthorNumber - j) / iSum
            # 使用get()函数返回特定键的值
            fAuthorCurrentWeight = dAuthorDict.get(sAuthorTemp, 0)
            # 值相加
            dAuthorDict[sAuthorTemp] = fAuthorCurrentWeight + fAuthorWeight
    
    # 将字典转为DataFrame
    dataframe_modified = pandas.DataFrame(dAuthorDict, index=[0])
    # print(dataframe_modified)
    
    # 由于字典转成的DataFrame为1行n列，因此进行转置
    dataframe_modified_T = dataframe_modified.T
    # print(dataframe_modified_T)
    
    # 将转置后的DataFrame按照加权统计降序排列
    # 参数解释:
    # axis: 0-纵向排序 1-横向排序
    # by: 如果axis=0，那么by="列名"；如果axis=1，那么by="行名"
    # ascending: 布尔型，True=升序，False=降序
    dataframe_modified_T_sorted = dataframe_modified_T.sort_values(axis=0, by=[0], ascending=False)
    
    # +---------------End---------------+
    
    # 写入Excel
    write = pandas.ExcelWriter(outfile_path)
    dataframe_modified_T_sorted.to_excel(write)
    write.save()
    
    print("完成")
