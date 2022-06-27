try{
#Specify the path to the Excel file and the WorkSheet Name
$FilePath = "C:\eTaskAutomationTesting\ImportData.xlsx"
$updateSheet = "Update"
$resultSheet = "Result"
#Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application
#Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false

#Set Display alerts as false
$objExcel.displayalerts = $False

#Open the Excel file and save it in $WorkBook
$WorkBook = $objExcel.Workbooks.Open($FilePath)
#Load the WorkSheet 'BuildSpecs'
$WorkSheet1 = $WorkBook.sheets.item($updateSheet)
$WorkSheet2 = $WorkBook.sheets.item($resultSheet)

#Deleting the worksheet
$WorkSheet1.Delete()
$WorkSheet2.Delete()
#Saving the worksheet
$WorkBook.Save()
$WorkBook.close($true)
$objExcel.Quit()
Write-Host "Result & Update sheets in ImportData.xlsx Successfully Deleted." -ForegroundColor Green
}
catch{
    Write-Host "Result & Update sheets in ImportData.xlsx Failed to delete due to sheets non-existent." -ForegroundColor Red
}