from newaccess import accesstoexcel, macro
from openpyxl import load_workbook

mix_type = "P-52"
title = "%s_Limits" % (mix_type)

sql = "SELECT * FROM MATERIAL_TABLE WHERE MATERIAL = '%s'"% (mix_type)
new_sql = "SELECT * FROM LIMITS_TABLE WHERE MATERIAL = '%s'"% (mix_type)
file_name = "Materials\\analysis(%s).MAC" % (mix_type.lower())
save_name1 = "Materials\\%s.xlsx" % (mix_type)
save_name2 = "Materials\\%s.xlsx" % (title)

accesstoexcel(save_name1,sql)
accesstoexcel(save_name2,new_sql)
macro(file_name,save_name2,sql,mix_type)
