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
    # 统计一共有多少行(文献篇数)
    nrows = dataframe.shape[0]
    # print(nrows)
    
    all_author_number = list(dataframe.count(axis=1))
    # print(all_author_number)
    # 第n篇文献的作者数为author_number[n-1]
    
    # 创建一个空字典
    dAuthorDict = {}
    
    # 开始定位单元格，按照单元格进行处理
    array_2d = dataframe.values
    
    for row in range(0, nrows):
        # 获得第row+1篇文献的作者数量
        author_number = all_author_number[row]
        # 计算分母
        total_weight = int(author_number * (author_number + 1) / 2)
        
        # 从第一个作者开始遍历
        for author in range(0, author_number):
            # 取单元格内的作者
            author_name = array_2d[row][author]
            # 计算单个作者的权重
            author_weight = (author_number - author) / total_weight
            # 使用get()函数返回特定键的值
            author_current_weight = dAuthorDict.get(author_name, 0)
            # 值相加
            dAuthorDict[author_name] = author_current_weight + author_weight
    
    # 将字典转为DataFrame
    dataframe_modified = pandas.DataFrame(dAuthorDict, index=["加权统计"])
    # print(dataframe_modified)
    
    # 由于字典转成的DataFrame为1行n列，因此进行转置
    dataframe_modified_T = dataframe_modified.T
    # print(dataframe_modified_T)
    
    # 将转置后的DataFrame按照加权统计降序排列
    # 参数解释:
    # axis: 0-纵向排序 1-横向排序
    # by: 如果axis=0，那么by="列名"；如果axis=1，那么by="行名"
    # ascending: 布尔型，True=升序，False=降序
    dataframe_modified_T_sorted = dataframe_modified_T.sort_values(axis=0, by=["加权统计"], ascending=False)
    
    # +---------------End---------------+
    
    # 写入Excel
    write = pandas.ExcelWriter(outfile_path)
    dataframe_modified_T_sorted.to_excel(write)
    write.save()
    
    print("完成")
