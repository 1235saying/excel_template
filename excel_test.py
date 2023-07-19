# -*- encoding: utf-8 -*-
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
#	@FileName   :	excel_test.py
#	@Time       :	2023/07/17 06:23:39
#	@Desc       :   测试excel_template.py
#	@Version    :
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #


from excel_template import excel_template
import pandas as pd

def my_print(contents: list):
    for content in contents:
        print(content.center(40, '='))
    return


class excel_operation(excel_template):
    # 
    def rule_deal_excel(self) -> list or dict:
        super().rule_deal_excel()
        
        target_data=[]

        

        for index,row in self.task_sheet.iterrows():
            row_data=[]
            for col_value in self.task_sheet.columns:

                new_col_value=row[col_value]+1 if isinstance(row[col_value],int) else row[col_value]

                row_data.append(new_col_value)

            target_data.append(row_data)

        my_print([f"{self.file_name}处理完成！！！"])

        return target_data





def main():

    excel=excel_operation()
    data=excel.rule_deal_excel()
    excel.write_excel(data)

    pass


if __name__ == '__main__':
    main()
