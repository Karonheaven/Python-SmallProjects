# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    infile = "3-分割作者&关键词.xlsx"
    outfile = "4-统计作者&关键词&文献来源&年份.xlsx"
    
    # 获取源Excel路径(infile_path)，并将数据分别写入dataframe中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe_author = pandas.read_excel(infile_path, sheet_name=0, header=0)
    dataframe_keyword = pandas.read_excel(infile_path, sheet_name=1, header=0)
    dataframe_source = pandas.read_excel(infile_path, sheet_name=2, header=0)
    dataframe_year = pandas.read_excel(infile_path, sheet_name=3, header=0)
    
    # +---------------Begin---------------+
    
    # 统计出现的次数 实际上就是.stack().value_counts()
    author_count = dataframe_author.stack().value_counts()
    keyword_count = dataframe_keyword.stack().value_counts()
    source_count = dataframe_source.stack().value_counts()
    year_count = dataframe_year.stack().value_counts()
    
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
    
    print("完成")
