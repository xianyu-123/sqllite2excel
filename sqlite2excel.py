"""
create by wanggj
"""
import sqlite3
import pandas as pd

class Sqlite2Excel(object):
    def __init__(self, sqlfile, save_file):
        """
        Example:
            >>> from sqlite2excel import Sqlite2Excel
            >>> Sqlite2Excel("test.sqllite", "test.xlsx")
        Args:
            sqlfile:sqllite文件
            save_file:要转换的xlsx文件
        """
        self.infile = sqlfile
        self.save_file = save_file
        
        self.conn = sqlite3.connect(self.infile) # 连接数据库
        self.cursor = self.conn.cursor()
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")#获取其中所有表名
        self.table = self.cursor.fetchall() #获取其中所有的表
        self.__save()
    
    def __save(self):
        writer = pd.ExcelWriter(self.save_file) # excel写入器
        for tb in self.table:
            # 获取列名
            self.cursor.execute(f"PRAGMA table_info({tb[0]})") 
            columns = self.cursor.fetchall()
            column_names = [x[1] for x in columns]
            # 执行查询并获取数据
            curtb = self.cursor.execute("SELECT * FROM "+tb[0])
            rows = curtb.fetchall()
            # 将数据转换为DataFrame 
            df = pd.DataFrame(rows, columns=column_names)
            # 将DataFrame写入Excel工作簿的新工作表中  
            df.to_excel(writer, sheet_name=tb[0], index=False)
        # 保存并关闭
        writer.save()
        writer.close()
