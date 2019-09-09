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
endOfWells = 2152


for x in range(1,endOfWells):

  wellsSerial = sheet.cell_value(x,5)
  wellsAssetPrice = sheet.cell_value(x,6)

  try:
    DLLAssetPrice = sheet2.cell_value(x,19)
    DLLserial = sheet2.cell_value(x,22)
  except:
    continue

  for y in range(1, endOfXLSheet):
    testSerial = sheet3.cell_value(y,10)
    
    if testSerial == "":
      break
    if testSerial == wellsSerial:
      try:
        worksheet.write(y,0, testSerial)
        worksheet.write(y,1, wellsAssetPrice)
      except:
        continue
      break
    if testSerial == DLLserial:
      try:
        worksheet.write(y,0, testSerial)
        worksheet.write(y,1, DLLAssetPrice)
      except:
        continue
      break

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))

# Take assetLvlPrice from LeaseData
# binary search for the serial number In SherpaReport
# inject assetLvlPrice in asselLvlPrice of that row



