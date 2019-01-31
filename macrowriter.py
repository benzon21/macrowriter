from newaccess import access_to_csv, macro

mix_type = "P-52"
title = "{0}_Limits".format(mix_type)
save_directory = "home\\Material"

material_sql = "SELECT * FROM MATERIAL_DATABASE WHERE MATERIAL = '{0}'".format(mix_type)
limits_sql = "SELECT * FROM LIMITS_DATABASE WHERE MATERIAL = '{0}'".format(mix_type)
file_name = "{0}\\analysis({1}).MAC".format(save_directory,mix_type.lower())
material_data = "{0}\\{1}".format(save_directory,mix_type)
limits_data = "{0}\\{1}".format(save_directory,title)

access_to_csv(material_data,material_sql)
access_to_csv(limits_data,limits_sql)
macro(file_name,limits_sql,material_sql)
