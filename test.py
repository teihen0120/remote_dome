#モジュールのインポート
import openpyxl as px
from openpyxl.styles import PatternFill
import os
import glob
import datetime

#エクセルファイル読み込み
###############エクセルが入っているフォルダのパスに書き換えてください##################
xlfile_folder_path = r"D:\D_tokudome\D_desktop\result_analysis\*.xlsx"
##################################################################################

#ファイル取得
xlfile_path_list = glob.glob(xlfile_folder_path)

###############保存先のパスに書き換えてください(tyousa, resultのところが保存ファイル名になっています)###################
#調査表保存用のパス
save_tyousa_path = r"D:\D_tokudome\D_desktop\tyousa_test"
#結果保存用のパス
save_result_path = r"D:\D_tokudome\D_desktop\result_test"
##################################################################################################################

#フォルダ生成
os.makedirs(save_tyousa_path, exist_ok=True)
os.makedirs(save_result_path, exist_ok=True)

#色指定
def get_red(plus_red=0):
    hex_red = f"{100+plus_red:02X}{0:02X}{0:02X}"
    red = PatternFill(patternType="solid", fgColor=hex_red, bgColor=hex_red)
    return red
def get_green(plus_green=0):
    hex_green = f"{0:02X}{100+plus_green:02X}{0:02X}"
    green = PatternFill(patternType="solid", fgColor=hex_green, bgColor=hex_green)
    return green
def get_blue(plus_blue=0):
    hex_blue = f"{0:02X}{0:02X}{100+plus_blue:02X}"
    blue = PatternFill(patternType="solid", fgColor=hex_blue, bgColor=hex_blue)
    return blue

#空のワークブック作成　result書き込み用
wb_init = px.Workbook()
wb_init["Sheet"].title = "Sheet1"
sheet_init = wb_init["Sheet1"]

wb_uekae_write = px.Workbook()
wb_uekae_write["Sheet"].title = "Sheet1"
sheet_uekae_write = wb_uekae_write["Sheet1"]

wb_pre = px.Workbook()
wb_pre["Sheet"].title = "Sheet1"
sheet_pre = wb_pre["Sheet1"]

wb_init_2 = px.Workbook()
wb_init_2["Sheet"].title = "Sheet1"
sheet_init_2 = wb_init_2["Sheet1"]

#出蕾日と開花日でブックを分割                    
wb_tsubomi = px.Workbook()
wb_tsubomi["Sheet"].title = "Sheet1"
sheet_tsubomi = wb_tsubomi["Sheet1"]

wb_flower = px.Workbook()
wb_flower["Sheet"].title = "Sheet1"
sheet_flower = wb_flower["Sheet1"]

wb_uekae = px.Workbook()
wb_uekae["Sheet"].title = "Sheet1"
sheet_uekae = wb_uekae["Sheet1"]

wb_tsubomi_2 = px.Workbook()
wb_tsubomi_2["Sheet"].title = "Sheet1"
sheet_tsubomi_2 = wb_tsubomi_2["Sheet1"]

wb_flower_2 = px.Workbook()
wb_flower_2["Sheet"].title = "Sheet1"
sheet_flower_2 = wb_flower_2["Sheet1"]

#出蕾日、開花日記録、色塗り
for n in range(0, len(xlfile_path_list)):
    print(xlfile_path_list[n])
    #現在と１つ前のワークブック作成
    wb = px.load_workbook(xlfile_path_list[n])
    sheet = wb["Sheet1"]
    
    #sheetの日にち取得
    date_name = os.path.splitext(os.path.basename(xlfile_path_list[n]))[0]
    year = int(date_name[0:4])
    month = int(date_name[4:6])
    day = int(date_name[6:8])
    #取得した日にちに１日足す
    next_date_name = datetime.date(year, month, day) + datetime.timedelta(days=1)
    next_date_name = next_date_name.strftime("%Y%m%d")
    
    #いらない栽培ベッドデータ削除
    for column in range(1, sheet.max_column):
        value = sheet.cell(row=1, column=column).value
        if value not in (4, 10, 16, 22, 27, None):
            sheet.delete_cols(column, 2)
    
    #初日のデータ        
    if n == 0:
        for row in range(1, sheet.max_row+1):
            for column in range(1, sheet.max_column+1):
                #出蕾日、開花日の欄の範囲のとき
                if (3 <= row) and (2 <= column):
                    sheet.cell(row=row, column=column).fill = get_blue()
                    
                    if sheet.cell(row=row, column=column).has_style:
                        sheet_pre.cell(row=row, column=column).fill = sheet.cell(row=row, column=column).fill._StyleProxy__target
                
                
    #２日目以降のデータ                        
    else:
        for row in range(3, sheet.max_row+1):
            for column in range(2, sheet.max_column+1):
                if sheet_pre.cell(row=row, column=column).fill == get_blue():
                    sheet.cell(row=row, column=column).fill = get_red()
                    # sheet.cell(row=row, column=column).fill = sheet_pre.cell(row=row, column=column).fill._StyleProxy__target
    
    for row in range(3, sheet.max_row+1):
        for column in range(2, sheet.max_column+1):
            sheet_pre.cell(row=row, column=column).fill = sheet.cell(row=row, column=column).fill._StyleProxy__target
    
                        
    #保存_調査票
    wb.save(save_tyousa_path + "\\" + str(next_date_name) + ".xlsx")

#保存_結果表（出力日は、エクセルリストの最後の日付）
wb_init.save(save_result_path + "\\" + str(date_name) + "_result.xlsx")
wb_init_2.save(save_result_path + "\\" + str(date_name) + "_result_02.xlsx")
wb_tsubomi.save(save_result_path + "\\" + str(date_name) + "_result_tsubomi.xlsx")
wb_flower.save(save_result_path + "\\" + str(date_name) + "_result_flower.xlsx")
wb_tsubomi_2.save(save_result_path + "\\" + str(date_name) + "_result_tsubomi_02.xlsx")
wb_flower_2.save(save_result_path + "\\" + str(date_name) + "_result_flower_02.xlsx")
wb_uekae.save(save_result_path + "\\" + str(date_name) + "_result_uekae.xlsx")