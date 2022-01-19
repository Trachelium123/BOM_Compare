import os
import pandas as pd
import numpy as np
import xlwt

f = xlwt.Workbook()
f1 = xlwt.Workbook()
# 1是旧的BOM，2是新的BOM，顺序为文件列数是['Quantity' , 'Reference' , 'PART_NUMBER']
file_path_old = 'old.xlsx'
file_path_new = 'new.xlsx'
# file_name_old = os.path.basename(file_path_old)
# file_old = pd.ExcelFile(file_path_old)

# 获取数据
file_old = pd.read_excel(file_path_old, header=0)
data_old = np.array(file_old)

file_new = pd.read_excel(file_path_new, header=0)
data_new = np.array(file_new)

partnum_old = []
partnum_new = []
m = 1
# 创建新表格
sheet1 = f.add_sheet(sheetname='111', cell_overwrite_ok=True)

# 删除料号
for t in range(data_old.shape[0]):
    PART_NUMBER_old1 = data_old[t][2]
    for j in range(data_new.shape[0]):
        partnum_new.append(data_new[j][2])
    if PART_NUMBER_old1 not in partnum_new:
        # print(str(m) + '、删除料号' + str(PART_NUMBER_old1) +'【中文说明】' + '\n变更原因：')
        sheet1.write(m, 0, str(m) + '、删除料号' + str(PART_NUMBER_old1) +'【中文说明】' + '\n变更原因：')
        m += 1
# 新增料号
for i in range(data_new.shape[0]):
    PART_NUMBER_new = data_new[i][2]
    # print(PART_NUMBER_new)
    for j in range(data_old.shape[0]):
        partnum_old.append(data_old[j][2])
    # print(partnum_old)
    if PART_NUMBER_new not in partnum_old:
        # print(str(m) + '、新增料号' + str(PART_NUMBER_new) +'【中文说明】，位号' + str(data_new[i][1]) + '，最终数量：'
        #       + str(data_new[i][0]) + '\n变更原因：')
        sheet1.write(m, 0, str(m) + '、新增料号' + str(PART_NUMBER_new) +'【中文说明】，位号' + str(data_new[i][1]) + '，最终数量：'
              + str(data_new[i][0]) + '\n变更原因：')
        m += 1
# 变更料号
for i in range(data_new.shape[0]):
    PART_NUMBER_new = data_new[i][2]
    # print(PART_NUMBER_new)
    for j in range(data_old.shape[0]):
        if data_old[j][2] == PART_NUMBER_new:
            PART_NUMBER_old = data_old[j][2]
            tem_new = data_new[i][1].split(',')
            tem_old = data_old[j][1].split(',')
            # print(len(tem_old),len(tem_new))
            tem_write_delete = []
            tem_write_add = []
            # 删除某东西
            for k in range(len(tem_old)):
                if tem_old[k] not in tem_new:
                    # print(tem_old[k])
                    tem_write_delete.append(tem_old[k])
            # 增加某东西
            for p in range(len(tem_new)):
                if tem_new[p] not in tem_old:
                    # print(tem_new[k])
                    tem_write_add.append(tem_new[p])

            if tem_write_delete and tem_write_add:
                # print(str(m)  + '、在料号' + PART_NUMBER_old +'【中文说明】中删除位号' + str(tem_write_delete) + '，' +
                #              '增加位号' + str(tem_write_add) + '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                # # print(tem_write_add)

                sheet1.write(m, 0, str(m) + '、在料号' + PART_NUMBER_old +'【中文说明】中，删除位号' + str(tem_write_delete) +
                             '，' + '增加位号' + str(tem_write_add) + '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                m += 1
            elif tem_write_delete and tem_write_add == []:
                # print(str(m)  + '、在料号' + PART_NUMBER_old +'【中文说明】中删除位号' + str(tem_write_delete) +
                #              '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                # # print(tem_write_add)

                sheet1.write(m, 0, str(m) + '、在料号' + PART_NUMBER_old +'【中文说明】中，删除位号' + str(tem_write_delete) +
                             '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                m += 1
            elif tem_write_delete == [] and tem_write_add:
                # print(str(m)  + '、在料号' + PART_NUMBER_old +'【中文说明】中' +
                #              '增加位号' + str(tem_write_add) + '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                # # print(tem_write_add)

                sheet1.write(m, 0, str(m)  + '、在料号' + PART_NUMBER_old +'【中文说明】中，' +
                             '增加位号' + str(tem_write_add) + '，最终数量：' + str(data_new[i][0]) + '\n变更原因：')
                m += 1
    # print(partnum_old)

# 保存临时表格

f.save('tem.xls')
# 替换无用字符
a = 0
file_path_tem = 'tem.xls'
file_tem = pd.read_excel(file_path_tem, header=None)
data_tem = np.array(file_tem)
sheet2 = f1.add_sheet(sheetname='zhu', cell_overwrite_ok=True)
for i in range(data_tem.shape[0]):
    data_tem[i][0] = str(data_tem[i][0]).replace('\'', '').replace('[', '').replace(']', '').replace('\"', '')
    sheet2.write(a, 0, data_tem[i][0])
    a += 1
f1.save('result.xls')
os.remove('tem.xls')
print('OK')
