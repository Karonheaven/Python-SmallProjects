# coding:utf-8
import sys
import pandas

if __name__ == "__main__":
    path = sys.path[0]
    # 或者path="."
    
    # 作者非空
    infile = "1-源数据.xlsx"
    outfile = "2-作者非空.xlsx"
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe = pandas.read_excel(infile_path, sheet_name=0, header=0)
    dataframe_modified = dataframe.dropna(subset=["Author-作者"])
    
    write = pandas.ExcelWriter(outfile_path)
    dataframe_modified.to_excel(write, index=False, engine="xlsxwriter")
    write.save()
    print("作者非空 完成")
    
    # 分割作者&关键词
    infile = "2-作者非空.xlsx"
    outfile = "3-分割作者&关键词.xlsx"
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe = pandas.read_excel(infile_path, sheet_name=0)
    dataframe_author = dataframe["Author-作者"]
    dataframe_author_modified = dataframe_author.str.split(";;|;|；|,", expand=True)
    dataframe_keyword = dataframe["Keyword-关键词"]
    dataframe_keyword_modified = dataframe_keyword.str.split(";;|;|；|,", expand=True)
    dataframe_source = dataframe["Source-文献来源"]
    dataframe_year = dataframe["Year-年"]
    
    write = pandas.ExcelWriter(outfile_path)
    dataframe_author_modified.to_excel(write, sheet_name="作者分割", index=False)
    dataframe_keyword_modified.to_excel(write, sheet_name="关键词分割", index=False)
    dataframe_source.to_excel(write, sheet_name="文献来源", index=False)
    dataframe_year.to_excel(write, sheet_name="年", index=False)
    write.save()
    print("分割 完成")
    
    # 统计作者&关键词&年份
    infile = "3-分割作者&关键词.xlsx"
    outfile = "4-统计作者&关键词&文献来源&年份.xlsx"
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe_author = pandas.read_excel(infile_path, sheet_name=0, header=0)
    dataframe_keyword = pandas.read_excel(infile_path, sheet_name=1, header=0)
    dataframe_source = pandas.read_excel(infile_path, sheet_name=2, header=0)
    dataframe_year = pandas.read_excel(infile_path, sheet_name=3, header=0)
    
    author_count = dataframe_author.stack().value_counts()
    keyword_count = dataframe_keyword.stack().value_counts()
    source_count = dataframe_source.stack().value_counts()
    year_count = dataframe_year.stack().value_counts()
    
    write = pandas.ExcelWriter(outfile_path)
    author_count.to_excel(write, sheet_name="作者统计")
    keyword_count.to_excel(write, sheet_name="关键词统计")
    source_count.to_excel(write, sheet_name="文献来源统计")
    year_count.to_excel(write, sheet_name="年份统计")
    write.save()
    print("统计 完成")
    
    # 作者加权统计
    infile = "3-分割作者&关键词.xlsx"
    outfile = "5-作者加权统计.xlsx"
    infile_path = path + "/" + infile
    outfile_path = path + "/" + outfile
    
    dataframe = pandas.read_excel(infile_path, sheet_name=0, header=0)
    nrows = dataframe.shape[0]
    all_author_number = list(dataframe.count(axis=1))
    dAuthorDict = {}
    array_2d = dataframe.values
    for row in range(0, nrows):
        author_number = all_author_number[row]
        total_weight = int(author_number * (author_number + 1) / 2)
        for author in range(0, author_number):
            author_name = array_2d[row][author]
            author_weight = (author_number - author) / total_weight
            author_current_weight = dAuthorDict.get(author_name, 0)
            dAuthorDict[author_name] = author_current_weight + author_weight
    
    dataframe_modified = pandas.DataFrame(dAuthorDict, index=[0])
    dataframe_modified_T_sorted = dataframe_modified.T.sort_values(axis=0, by=[0], ascending=False)
    
    write = pandas.ExcelWriter(outfile_path)
    dataframe_modified_T_sorted.to_excel(write)
    write.save()
    print("加权统计 完成")
