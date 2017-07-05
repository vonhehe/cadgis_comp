#-*-coding:utf-8-*-
import xlrd
import xdrlib ,sys
import os

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

def gis_table_byindex(file,by_index=0):
    gis_col = []
    data = open_excel(file)
    table = data.sheets()[by_index]  #通过索引
    #nrows = table.nrows  #行数
    #ncols = table.ncols  #列数

    n=0     #标记用地面积的位置
    for j in table.row_values(0):
        if j == "DKMJ":
            break
        else:
            n += 1
            
    gis_col = table.col_values(n)
    del gis_col[0]   #表头删掉
    #hang = table.row(6)[14].value
    #print(gis_col)
    return gis_col

 
def cad_table_byindex(file,by_index=0):
    cad_col = []
    data = open_excel(file)
    table = data.sheets()[by_index]  #通过索引
    nrows = table.nrows  #行数

    n=0     #标记用地面积的位置
    for j in table.row_values(0):
        if j == "FD_YDMJ":
            break
        else:
            n+=1
    #ncols = table.ncols  #列数
    
    #把里表里的字符串转成数字型，先删除前两个文字和字母
    cad_col = table.col_values(n)
    del cad_col[0]
    del cad_col[0]  
    cad_col = [float(i) for i in cad_col]    

    #hang = table.row(6)[14].value
    #print(cad_col)
    #print(type(cad_col[1]))
    return cad_col

def compare(file,gis,cad,by_index=0):#比较数值，由于cad导出的加了一个合计，所以cad里的长度应为 len(gis)=len(cad)-1
    cad_dkbh = ""
    rent = 0  #地块面积的差值
    data = open_excel(file)
    table = data.sheets()[by_index]
    cad_dkbh = table.col_values(1)

    #for cad_bh in cad_dkbh:
        #for gis_bh in 
    
    print("{0:10}\t{1}".format("不符合要求的地块编号","相差面积"))
    #for bh in 
    for i in range(len(gis)):
        rent = abs(gis[i]-cad[i])
        cad_dkbh = table.row(i+2)[1].value
        if rent>0.5:
            print("{0:20}\t{1}".format(cad_dkbh,rent))
        else:
            continue
            
            
def main():
    #file1 = 'Output.xlsx'
    file1 = r'Output.xlsx'
    file2 = r'file2.xls'
    #gis_col = 15
    #cad_col = 4
    
    gis = gis_table_byindex(file1)
    cad = cad_table_byindex(file2)

    compare(file2,gis,cad)
    os.system("pause")
main()
