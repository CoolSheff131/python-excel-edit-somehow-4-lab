from openpyxl import load_workbook
excel_file = './1.xlsx'
wb = load_workbook(excel_file, data_only = True)
sh = wb['Sheet1']
color_in_hex = sh['A2'].fill.start_color.index # this gives you Hexadecimal value of the color
print ('HEX =',color_in_hex) 
color_in_hex = sh['B2'].fill.start_color.index # this gives you Hexadecimal value of the color
print ('HEX =',color_in_hex)
color_in_hex = sh['C2'].fill.start_color.index # this gives you Hexadecimal value of the color
print ('HEX =',color_in_hex)
color_in_hex = sh['D2'].fill.start_color.index # this gives you Hexadecimal value of the color
print ('HEX =',color_in_hex)