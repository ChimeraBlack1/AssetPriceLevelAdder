import math
import xlrd
import xlwt

#open wells workbook
loc = ("WellsPortfolio.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#open DLL workbook
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
endOfDLL = 1380


for x in range(1,endOfXLSheet):
  # get serial to test
  try:
    testSerial = sheet3.cell_value(x,10)
  except:
    continue

  # look in the wells portfolio for the serial
  for y in range(1, endOfWells):
    try:
      wellsSerial = sheet.cell_value(y,5)
      wellsAssetPrice = sheet.cell_value(y,6)
    except:
      continue

    if testSerial == "":
      break
    if testSerial == wellsSerial:
      try:
        worksheet.write(x,0, testSerial)
        worksheet.write(x,1, wellsAssetPrice)
      except:
        continue
      break

  # look in the DLL portfolio for the serial
  for y in range(1, endOfDLL):
    try:
      DLLAssetPrice = sheet2.cell_value(y,19)
      DLLserial = sheet2.cell_value(y,22)
    except:
      continue

    if testSerial == "":
      break
    if testSerial == DLLserial:
      try:
        worksheet.write(x,0, testSerial)
        worksheet.write(x,1, DLLAssetPrice)
      except:
        continue
      break


workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))