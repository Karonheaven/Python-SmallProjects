# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    infile = "a.xlsx"
    outfile = "b.xlsx"

    # 生成源Excel路径(infile_path)和目标Excel路径(outfile_path)
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    #
    # infile_path="./"+infile
    
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
    
    print("完成")
