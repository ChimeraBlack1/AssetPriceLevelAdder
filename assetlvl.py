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
          #return true if found
          found = True
        else:
          if theList[arrVal] < target:
            xmin = arrVal + 1
          else:
            xmax = arrVal - 1
    #return false if not found
    return found


loc = ("reportForPerry09062019.xlsm")
wb = xlrd.open_workbook(loc)

#open workbook
sheet = wb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Unique Leases')

# workbook.save(NewWorkbookName)

# Take leaseNumber, and the assetLvlPrice from LeaseData
# Find lease number In SherpaReport
  # Then Take serial number from LeaseData
  # Then Find serial number In SherpaReport
    # Take row number in SherpaReport

  # if leaseNumber == sherpaLeaseNumber and serialNumber == sherpaSerialNumber:
      # write the assetLvlPrice in assetLvlPrice column of THAT row
