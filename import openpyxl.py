#hiiiiiiiiiiiiiiiiiiiiiiiiii
we are on the develop branch

add a new feature is greate


testing

wb1 = openpyxl.load_workbook('sheet1.xlsx')
sheet1 = wb1.worksheets[0]
ws1 = wb1.active

row_count1 = sheet1.max_row
column_count1 = sheet1.max_column

#print(row_count1)
#print(column_count1)

wb2 = openpyxl.load_workbook('sheet2.xlsx')
sheet2 = wb2.worksheets[0]
ws2 = wb2.active

row_count2 = sheet2.max_row
column_count2 = sheet2.max_column



book = Workbook()
sheet = book.active

sheet.column_dimensions['A'].width = 60
sheet.column_dimensions['B'].width = 45
sheet.column_dimensions['C'].width = 45
sheet.column_dimensions['D'].width = 25

tuple(ws1.rows)

sheet['A1'] = 'Author Name'
sheet['A1'].fill = PatternFill(start_color="FFBD33", end_color="FFBD33", fill_type = "solid")

sheet['B1'] = 'Location in sheet 1'
sheet['B1'].fill = PatternFill(start_color="FFBD33", end_color="FFBD33", fill_type = "solid")

sheet['C1'] = 'Location in sheet 2'
sheet['C1'].fill = PatternFill(start_color="FFBD33", end_color="FFBD33", fill_type = "solid")


sheet['D1'] = 'Total Duplications'
sheet['D1'].fill = PatternFill(start_color="FFBD33", end_color="FFBD33", fill_type = "solid")


#redFill = PatternFill(color='FFFF0000')

#sheet['A1'].style = redFill
#sheet['B1'].style = redFill
#sheet['C1'].style = redFill



#print(row_count2)
#print(column_count2)


x = 1  
y = 1
z = 1
w = 2

for col in sheet1['E2:E5590']:

        for cell in col:

                    for col2 in sheet2['D2:D457']:
                    
                        for cell2 in col2:
                            
                            if cell2.value == cell.value:
                                
                                    x = x + 1 
                                    y = y + 1
                                    z = z + 1
                                   
                                    
                                
                                
                                    print ("--------------------------------------------------")
                                    print(cell2.value)

                                    print(cell.coordinate)

                                    print(cell2.coordinate)

                                    print ("-----------------------------------------------------")
                                    
                                    
                                    
                                    sheet.cell(row=x, column=1).value = cell2.value

                                    sheet.cell(row=y, column=2).value = cell.coordinate
                                    
                                    sheet.cell(row=z, column=3).value = cell2.coordinate

                                     ###########
                                        
                                    sheet.cell(row=w, column=4).value = x 
                                    
                                    cell.fill = PatternFill(start_color="008000", end_color="008000", fill_type = "solid")
                                    cell2.fill = PatternFill(start_color="008000", end_color="008000", fill_type = "solid")
                                    cell.font = Font(color='FFFFFF')
                                    cell2.font = Font(color='FFFFFF')
    
                                    
sheet.delete_rows(137,592) 

x = x-1 
    
book.save('author_result.xlsx') 

wb1.save('sheet1.xlsx') 
wb2.save('sheet2.xlsx') 


print ('Total duplicated')                                
print (sheet.max_row)   
              
                
