import math
import xlrd
import xlwt

def BinarySearch(theList, xmin, xmax, target):
    """
     This search method is the binary search method. It is a common algorithm for finding a target in a list where the list is already sorted either numerically or alphabetically.
    """
    found = False
    while xmin <= xmax and not found:
        arrVal = math.floor((xmin + xmax) / 2)
        if theList[arrVal] == target:
          found = True
        else:
          if theList[arrVal] < target:
            xmin = arrVal + 1
          else:
            xmax = arrVal - 1
    return found

loc = ("reportForPerry09062019.xlsm")
wb = xlrd.open_workbook(loc)

#open workbook
sheet = wb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Unique Leases')

LeaseList = []
testMin = 0
testMax = len(LeaseList) - 1
target = 67
iteration = 0
NewWorkbookName = "unique leases yo.xls"

excelSize = 10

for x in range(1,excelSize):
  leaseNumber = sheet.cell_value(x,3)
  # if the LeaseList is empty, add the leaseNumber to it
  if len(LeaseList) == 0:
    print(str(leaseNumber))
    LeaseList.append(leaseNumber)
    worksheet.write(x, 0, leaseNumber)
  elif len(LeaseList) > 1:
    # else if the len(LeaseList) > 1, run the binary search on that list for the leaseNumber
    LNIsFound = BinarySearch(LeaseList, testMin, testMax, leaseNumber)
    # if leaseNumber is NOT found, write to excel
    if LNIsFound == False:


      worksheet.write(x, 4, leaseNumber)
      print("wrote: " + str(leaseNumber) + " to excel row: " + str(x))

    #if leaseNumber IS found, jump to the next iteration
    if LNIsFound:
      continue
    
  else:
    # else add leaseNumber to the list
    LeaseList.append(leaseNumber)
    worksheet.write(x, 0, leaseNumber)

workbook.save(NewWorkbookName)



    
