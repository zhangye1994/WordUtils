import os

import win32com.client as win32


class Reader:
    doc = None
    word = None

    def __init__(self, file_path):
        # 创建或连接到Word应用程序实例
        word = win32.Dispatch('Word.Application')
        # 显示Word窗口（可选）
        word.Visible = False
        # 打开一个已存在的文档
        if not os.path.exists(file_path):
            raise Exception(f"{file_path} not exists")
        self.doc = word.Documents.Open(file_path)
        self.word = word

    def read_table_data(self, index):

        # 获取第一个表格（假设文档中有至少一个表格）
        tables = self.doc.Tables
        table = tables(index)
        data_list = []
        # 遍历表格的行和列
        for row in range(table.Rows.Count):
            data = []
            for col in range(table.Columns.Count):
                cell = table.Cell(row + 1, col + 1)  # 行索引和列索引从1开始
                data.append(cell.Range.Text)
            data_list.append(data)
        return data_list

    def read_paragraph(self):
        data_list = []
        # 遍历文档的所有段落并打印文本
        for paragraph in self.doc.Paragraphs:
            data_list.append(paragraph)
        return data_list

    def close(self):
        """
        关闭文档

        :return: None
        """
        # 关闭文档
        self.doc.Close()
        # 退出Word应用程序
        self.word.Quit()
