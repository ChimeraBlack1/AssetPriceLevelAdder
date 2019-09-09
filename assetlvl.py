import math
import xlrd
import xlwt

#open wells workbook
loc = ("WellsPortfolio.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#open wells workbook
loc2 = ("DLLPortfolio.xlsx")
wb2 = xlrd.open_workbook(loc2)
sheet2 = wb2.sheet_by_index(0)

#open report for perry
loc3 = ("reportForPerry09062019.xlsm")
wb3 = xlrd.open_workbook(loc3)
sheet3 = wb3.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('AssetLevelAdded')
NewWorkbookName = "myNewWb.xls"

# end of excel sheet (report for perry)
endOfXLSheet = 3587

for x in range(1,10):


  
  # wellsSerial = sheet.cell_value(x,5)
  # wellsAssetPrice = sheet.cell_value(x,6)
  # DLLAssetPrice = sheet2.cell_value(x,19)
  # DLLserial = sheet2.cell_value(x,22)

  for y in range(1, endOfXLSheet):
    testSerial = sheet3.cell_value(y,10)
    if testSerial == wellsSerial:
      worksheet.write(x,1, testSerial)
      worksheet.write(x,2, wellsAssetPrice)
      break
    if testSerial == DLLserial:
      worksheet.write(x,4, testSerial)
      worksheet.write(x,5, DLLAssetPrice)
      break

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))

# Take assetLvlPrice from LeaseData
# binary search for the serial number In SherpaReport
# inject assetLvlPrice in asselLvlPrice of that row



