# -*- encoding: utf-8 -*-
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
#	@FileName   :	excel_template.py
#	@Time       :	2023/07/17 04:21:28
#	@Desc       :   1、一个excel 读写模板
#                   2、实现了一个sheet 页 的 读取和写入（a\w）
#	@Version    :   1.0
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #

import os
import sys
import yaml
from openpyxl import load_workbook
import pandas as pd


yaml_file_name="excel.yaml"


def my_print(contents: list):
    for content in contents:
        print(content.center(40, '='))
    return


def file_exists(file_name):
    if not os.path.exists(file_name):
        my_print([f"{file_name}文件不存在！！！"])
        sys.exit(2)


class excel_template:
    def __init__(self):
        
        print(os.getcwd())
        # form yaml file get config
        file_exists(yaml_file_name)
        with open(yaml_file_name, 'r', encoding='utf-8') as f:
            self.data=yaml.safe_load(f)

        # config messages
        self.file_name=self.data.get("file_name",None)
        if not self.file_name:
            my_print([f"{self.file_name} is None"])
            sys.exit(2)
        file_exists(self.file_name)

        self.sheet_name=self.data.get("sheet_name","Sheet1") or "Sheet1"

        self.header=self.data.get("header",None)
        if not self.header:
            self.header = self.header if isinstance(self.header,int) else None

        self.usecols=self.data.get("usecols",None) or None
        if not self.usecols:
            self.usecols = self.usecols if isinstance(self.usecols,int) else None

        self.writer_mode=self.data.get("writer_mode","w")

        self.task_file_name=self.data.get("task_file_name",f"ok_{self.file_name}") or f"ok_{self.file_name}"
        if self.task_file_name is None:
            my_print([f"{self.task_file_name} is None"])
            sys.exit(2)

        self.task_sheet_name=self.data.get("task_sheet_name",self.sheet_name) or self.sheet_name
        if self.task_sheet_name is None:
            my_print([f"{self.task_sheet_name} is None"])
            sys.exit(2)        
        

    def read_excel(self):
        # read excel file
        # return a dataframe

        return pd.read_excel(
            self.file_name,
            sheet_name=self.sheet_name,
            header=self.header,
            usecols=self.usecols,
        )


    def rule_deal_excel(self) -> list or dict:
        self.task_sheet=self.read_excel()
        
        my_print([f"{self.file_name}:{self.sheet_name}读取完成······"])
    

    def write_excel(self,data:list or dict):
        # write excel file
        # return a dataframe
        df=pd.DataFrame(data)

        if self.writer_mode=="a":
            self._extracted_from_write_excel_a(df)
        elif self.writer_mode=="w":
            self._extracted_from_write_excel_w(df)

        my_print(["excel文件写入成功"])


        pass

    # writer_mode a
    def _extracted_from_write_excel_a(self, df):
        # 在 模板 excel 中写入
        file_exists(self.task_file_name)

        book=load_workbook(self.task_file_name)
        sheet=book[self.task_sheet_name]

        for index,row in df.iterrows():
            sheet.append(list(row))

        book.save(self.task_file_name)
        book.close()



    # writer_mode w
    def _extracted_from_write_excel_w(self,df):
        with pd.ExcelWriter(self.task_file_name, engine='openpyxl') as writer:
            df.to_excel(
                writer, 
                sheet_name=self.task_sheet_name, 
                index=False,
                header=False
            )
            





    



