# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    infile = "2-作者非空.xlsx"
    outfile = "3-分割作者&关键词.xlsx"
    
    # 获取源Excel路径(infile_path)，并将数据写入dataframe2中
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe = pandas.read_excel(infile_path, sheet_name=0)
    
    # +---------------Begin---------------+
    # 分割作者&关键词&单位并分别保存到新dataframe
    # 先将"Author-作者"的数据转换为Series，然后进行分割
    dataframe_author = dataframe["Author-作者"]
    dataframe_author_modified = dataframe_author.str.split(";;|;|；|,", expand=True)
    # 再将"Keyword-关键词"做相同的处理
    dataframe_keyword = dataframe["Keyword-关键词"]
    dataframe_keyword_modified = dataframe_keyword.str.split(";;|;|；|,", expand=True)
    
    # 文献来源Source和年份year不需要分割，但是为了module3方便统计，一并写入工作簿中
    dataframe_source = dataframe["Source-文献来源"]
    dataframe_year = dataframe["Year-年"]
    
    # +---------------End---------------+
    
    # 通过ExcelWriter将四个表一次性写入一个工作簿中
    # sheet1: 作者分割
    # sheet2: 关键词分割
    # sheet3: 文献来源
    # sheet4: 年
    write = pandas.ExcelWriter(outfile_path)
    
    dataframe_author_modified.to_excel(write, sheet_name="作者分割", index=False)
    dataframe_keyword_modified.to_excel(write, sheet_name="关键词分割", index=False)
    dataframe_source.to_excel(write, sheet_name="文献来源", index=False)
    dataframe_year.to_excel(write, sheet_name="年", index=False)
    
    write.save()
    
    print("完成")
