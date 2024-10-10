import sys
from time import sleep
import win32com.client
import os
import json
import re


import re


def matches_parentheses(s):
    # 匹配以左小括号开始和右小括号结束的字符串
    pattern = r"^\(.*$"
    return bool(re.match(pattern, s))


def is_number(s):
    # 匹配整数、浮点数、带有 '%' 的数值，并支持千分符
    pattern = r"^-?\d{1,3}(,\d{3})*(\.\d+)?%?$"
    return bool(re.match(pattern, s.strip()))


def write_file(data, file_path):
    """将嵌套列表写入到文件"""
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(
            data, f, ensure_ascii=False, indent=4
        )  # 使用 indent 参数以便格式化输出

    pass


def read_table_data(doc_path, file_name):
    # 创建Word应用程序
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 让Word在后台运行

    doc = None  # 初始化 doc 变量
    try:
        # 打开文档
        doc_path = os.path.join(os.getcwd(), doc_path)
        doc = word.Documents.Open(doc_path)
        table_data = []

        # 检查文档中是否有表格
        if doc.Tables.Count == 0:
            print(f"文档 {doc_path} 中没有表格")
            return []

        # 遍历文档中的表格
        for table_index, table in enumerate(doc.Tables):
            table_content = []
            # print(f"表格 {table_index + 1} 内容:")  # 打印表格索引
            # print(
            #     f"行数: {table.Rows.Count}, 列数: {table.Columns.Count}"
            # )  # 打印行数和列数
            for row in range(1, table.Rows.Count + 1):
                row_content = []
                for col in range(1, table.Columns.Count + 1):
                    try:
                        # 读取单元格内容并去掉多余的空格
                        cell_text = table.Cell(row, col).Range.Text
                        cell_text = cell_text.strip("\r\x07")
                        if cell_text == "":
                            continue
                        # print(cell_text)
                        if is_number(cell_text):
                            row_content.append(cell_text)
                        if matches_parentheses(cell_text):
                            cell_text = cell_text.replace(")", "")
                            cell_text = cell_text.replace("(", "-")
                            row_content.append(cell_text)
                    except Exception as e:
                        pass
                if len(row_content) > 0:
                    table_content.append(row_content)
                    # print("\t".join(row_content))  # 打印行内容
            if len(table_content) > 0:
                table_data.append(table_content)

    except Exception as e:
        print(f"打开文档时出错: {e}")
        return []

    finally:
        doc.Close(False)
        word.Quit()
    sleep(5)  # 避免还没关闭,打开下一个失败

    "写文件 "
    write_file(table_data, file_name)
    return table_data


def compare_tables(en_table, cn_table):

    for i in range(len(en_table)):
        for j in range(len(en_table[i])):
            if en_table[i][j] in cn_table[i]:
                continue
            else:
                print(f"在en中{en_table[i][j]}有异常,需要手动验证")
                break


if len(sys.argv) != 3:
    print(f"compare v1.0.1 by lhy")
    print(f"usage: compare en_doc_file cn_doc_file")
    sys.exit(0)


# 读取中英文文档中的表格数据
english_doc = sys.argv[1]
chinese_doc = sys.argv[2]

english_tables = read_table_data(english_doc, "en.txt")
chinese_tables = read_table_data(chinese_doc, "cn.txt")
# 对比两个表格是否一致
compare_tables(english_tables, chinese_tables)
