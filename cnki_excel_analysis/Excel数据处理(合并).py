# coding=utf-8
# +---------------+---------------+
# Name: 通过pandas处理Excel(多文件)
# +---------------+---------------+

# +---------------+---------------+
# 所需的包
# +---------------+---------------+
import pandas
import sys

# ---------------+---------------+
# module 0
# 功能：一个框架，写入写出
# ---------------+---------------+
"""
一个完整的文件路径=文件所在的文件夹路径+文件名
"""


def module0(infile, outfile, path):
    # 通过infile生成源Excel路径(infile_path)
    infile_path = path + "/" + infile
    # 通过outfile生成目标Excel路径(outfile_path)
    outfile_path = path + "/" + outfile
    
    # 读取Excel，将数据保存在dataframe中
    # 参数解释：
    # infile_path: 不解释，源Excel文件的路径
    # sheet_name: 选取读取Excel里面的哪张表(可以用索引也可以用表名)
    # header: 指定列名(None表示没有列名，0表示第一行作为列名，以此类推)
    #     如果不指定，则为None
    dataframe = pandas.read_excel(infile_path, sheet_name=0, header=None)
    
    # 对dataframe进行处理
    # <----------这里写dataframe的处理代码---------->
    
    dataframe_modified = dataframe
    
    # <----------这里写dataframe的处理代码---------->
    
    # 保存dataframe_modified到文件
    write = pandas.ExcelWriter(outfile_path)
    # 将dataframe_modified写入Excel，使用to_excel函数
    # 参数解释:
    # write: 暂时不需要动
    # sheet_name: 保存成Excel的表名
    # index: 是否把行名写进表(默认是True)
    # header: 把第几行列名写进去(默认是None，也就是不写列名)
    dataframe_modified.to_excel(write, sheet_name="Sheet1", index=False, header=None)
    write.save()
    
    print("Done")


# ---------------+---------------+
# module 1
# 功能：滤掉作者为空的非正式文献
# ---------------+---------------+
"""
# 参考材料：
https://blog.csdn.net/brink_compiling/article/details/76890198
python3 pandas读写excel

https://blog.csdn.net/brucewong0516/article/details/79097909
【python】pandas库pd.to_excel操作写入excel文件参数整理与实例
"""


def module1(infile, outfile, path):
    # 获取源Excel路径(infile_path)，并将数据写入dataframe1中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe1 = pandas.read_excel(infile_path, sheet_name=0, header=0)
    # 这里sheetname=0表示取索引为0的表(即第一个表)
    
    # +---------------Begin---------------+
    
    # 对dataframe1进行清洗
    dataframe1_modified = dataframe1.dropna(subset=["Author-作者"])
    
    # +---------------End---------------+
    
    # 保存dataframe1到新的文件
    write = pandas.ExcelWriter(outfile_path)
    
    dataframe1_modified.to_excel(write)
    write.save()
    print("作者非空 处理成功 文件已保存")


# ---------------+---------------+
# module 2
# 功能：分割作者&关键词
# ---------------+---------------+
"""
# 参考材料：
https://zhuanlan.zhihu.com/p/30529129
Python3 pandas库(15) 分列 （上）str.split()

https://blog.csdn.net/wangxingfan316/article/details/79628463
DataFrame.to_excel多次写入不同sheet
"""


def module2(infile, outfile, path):
    # 获取源Excel路径(infile_path)，并将数据写入dataframe2中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe2 = pandas.read_excel(infile_path, sheet_name=0)
    
    # +---------------Begin---------------+
    
    # 分割作者&关键词&单位并分别保存到新dataframe
    # 先将"Author-作者"的数据转换为Series，然后进行分割
    dataframe2_author = dataframe2["Author-作者"]
    dataframe2_author_modified = dataframe2_author.str.split(
        ";;|;|；|,", expand=True)
    # 再将"Keyword-关键词"做相同的处理
    dataframe2_keyword = dataframe2["Keyword-关键词"]
    dataframe2_keyword_modified = dataframe2_keyword.str.split(
        ";;|;|；|,", expand=True)
    
    # 文献来源Source和年份year不需要分割，但是为了module3方便统计，一并写入工作簿中
    dataframe2_source = dataframe2["Source-文献来源"]
    dataframe2_year = dataframe2["Year-年"]
    
    # +---------------End---------------+
    
    # 通过ExcelWriter将四个表一次性写入一个工作簿中
    # sheet1: 作者分割
    # sheet2: 关键词分割
    # sheet3: 文献来源
    # sheet4: 年
    write = pandas.ExcelWriter(outfile_path)
    
    dataframe2_author_modified.to_excel(write, sheet_name="作者分割", index=False)
    dataframe2_keyword_modified.to_excel(write, sheet_name="关键词分割", index=False)
    dataframe2_source.to_excel(write, sheet_name="文献来源", index=False)
    dataframe2_year.to_excel(write, sheet_name="年", index=False)
    
    write.save()
    
    print("作者分割 关键词分割 处理成功 文件已保存")


# ---------------+---------------+
# module 3
# 功能：统计作者&关键词&文献来源&年份
# ---------------+---------------+
def module3(infile, outfile, path):
    # 获取源Excel路径(infile_path)，并将数据分别写入dataframe3中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe3_author = pandas.read_excel(infile_path, sheet_name=0, header=0)
    dataframe3_keyword = pandas.read_excel(infile_path, sheet_name=1, header=0)
    dataframe3_source = pandas.read_excel(infile_path, sheet_name=2, header=0)
    dataframe3_year = pandas.read_excel(infile_path, sheet_name=3, header=0)
    
    # +---------------Begin---------------+
    
    # 统计出现的次数
    author_count = dataframe3_author.stack().value_counts()
    keyword_count = dataframe3_keyword.stack().value_counts()
    source_count = dataframe3_source.stack().value_counts()
    year_count = dataframe3_year.stack().value_counts()
    
    # +---------------End---------------+
    
    # 保存到新文件
    # sheet1: 作者统计
    # sheet2: 关键词统计
    # sheet3: 文献来源统计
    # sheet4: 年份统计
    write = pandas.ExcelWriter(outfile_path)
    
    author_count.to_excel(write, sheet_name="作者统计")
    keyword_count.to_excel(write, sheet_name="关键词统计")
    source_count.to_excel(write, sheet_name="文献来源统计")
    year_count.to_excel(write, sheet_name="年份统计")
    
    write.save()
    
    print("作者统计 关键词统计 文件来源统计 年份统计 处理成功 文件已保存")


# ---------------+---------------+
# module 4
# 功能：作者加权统计
# ---------------+---------------+
def module4(infile, outfile, path):
    # 获取源Excel路径(filepath)
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    # 利用pandas打开Excel并读取数据，所得数据写入dataframe4中
    dataframe4 = pandas.read_excel(infile_path, sheet_name=0, header=0)
    
    # +---------------Begin---------------+
    
    # 统计dataframe中一共有多少行(nrows)，多少列(ncolumns)
    # dataframe.shape返回一个tuple，第一个元素是行数，第二个元素是列数
    nrows = dataframe4.shape[0]
    ncolumns = dataframe4.shape[1]
    
    # 转换成二维数组，直接用基础语句解决
    # dataframe
    # 1, 2
    # 3, 4
    # array_2d=dataframe.values
    # array_2d
    # [[1, 2],
    #  [3, 4]]
    array_2d = dataframe4.values
    
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
            if (str(sAuthorTemp) == "nan"):
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
    
    # 前面已经计算完成作者的权重
    # 后面就是要把作者的权重这个字典写到Excel里
    # 两个步骤，第一个步骤是把字典转换成dataframe
    # 第二个步骤是把dataframe写到Excel里
    
    # 将字典转为DataFrame
    dataframe4_modified = pandas.DataFrame(dAuthorDict, index=[0])
    # print(dataframe4_modified)
    
    # 由于字典转成的DataFrame为1行n列，因此进行转置
    dataframe4_modified_T = dataframe4_modified.T
    # print(dataframe4_modified_T)
    
    # 将转置后的DataFrame按照加权统计降序排列
    # 参数解释:
    # axis: 0-纵向排序 1-横向排序
    # by: 如果axis=0，那么by="列名"；如果axis=1，那么by="行名"
    # ascending: 布尔型，True=升序，False=降序
    dataframe4_modified_T_sorted = dataframe4_modified_T.sort_values(
        axis=0, by=[0], ascending=False)
    
    # +---------------End---------------+
    
    # 写入Excel
    write = pandas.ExcelWriter(outfile_path)
    dataframe4_modified_T_sorted.to_excel(write)
    write.save()
    
    print("作者加权统计 处理成功 文件已保存")


# ---------------+---------------+
# 主函数
# ---------------+---------------+
if __name__ == "__main__":
    path = sys.path[0]
    
    # module0("0-test-infile.xlsx", "0-test-outfile.xlsx", path)
    
    module1("1-源数据.xlsx", "2-作者非空.xlsx", path)

    module2("2-作者非空.xlsx", "3-分割作者&关键词.xlsx", path)

    module3("3-分割作者&关键词.xlsx", "4-统计作者&关键词&文献来源&年份.xlsx", path)
    
    module4("3-分割作者&关键词.xlsx", "5-作者加权统计.xlsx", path)
