# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    infile = "1-源数据.xlsx"
    outfile = "2-作者非空.xlsx"
    
    # 获取源Excel路径(infile_path)，并将数据写入dataframe1中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe = pandas.read_excel(infile_path, sheet_name=0, header=0)
    # 这里sheetname=0表示取索引为0的表(即第一个表)
    
    # +---------------Begin---------------+
    
    # 对dataframe1进行清洗
    dataframe_modified = dataframe.dropna(subset=["Author-作者"])
    
    # +---------------End---------------+
    
    # 保存dataframe1到新的文件
    write = pandas.ExcelWriter(outfile_path)

    dataframe_modified.to_excel(write, index=False, header=True, engine="xlsxwriter")
    write.save()
    print("完成")
