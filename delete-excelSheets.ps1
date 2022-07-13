try {
    Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
}
catch {
    Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
}

try {
    #Specify the path to the Excel file and the WorkSheet Name
    $FilePath = "C:\eTaskAutomationTesting\ImportData.xlsx"
    $configSheet = "Config"

    $updateSheet = "Update"
    $updatecompareSheet = "Update (2)"
    $resultSheet = "Created-Result"
    $updateResultSheet = "Updated-Result"
    $summarySheet = "Data Summary"
    #Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    #Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false

    #Set Display alerts as false
    $objExcel.displayalerts = $False

    #Open the Excel file and save it in $WorkBook
    $WorkBook = $objExcel.Workbooks.Open($FilePath)
    #Load the WorkSheet 'BuildSpecs'

    $WorkSheet1 = $WorkBook.sheets.item($configSheet)
    
    #Deleting the worksheet
    if ($WorkSheet1) {
        $WorkSheet2 = $WorkBook.sheets.item($updateSheet)
        $WorkSheet3 = $WorkBook.sheets.item($resultSheet)
        $WorkSheet4 = $WorkBook.sheets.item($updatecompareSheet)
        $WorkSheet5 = $WorkBook.sheets.item($summarySheet)
        $WorkSheet2.Delete()
        $WorkSheet3.Delete()
        $WorkSheet4.Delete()
        $WorkSheet5.Delete()
        
    }
    else {
        $WorkSheet4 = $WorkBook.sheets.item($updateResultSheet)
        $WorkSheet1.Delete()
        $WorkSheet2.Delete()
        $WorkSheet3.Delete()
        $WorkSheet4.Delete()
        $WorkSheet5.Delete()
    }

    #Saving the worksheet
    $WorkBook.Save()
    $WorkBook.close($true)
    $objExcel.Quit()
    Write-Host "All sheets in ImportData.xlsx except for Config & Data-Import Successfully Deleted." -ForegroundColor Green
}
catch {
    Write-Host "Sheets in ImportData.xlsx Failed to delete due to sheets non-existent." -ForegroundColor Red
    $WorkBook.Save()
    $WorkBook.close($true)
    $objExcel.Quit()
}

try {
    Get-Process | Where-Object MainWindowTitle -eq 'ImportData.xlsx - Excel' | Stop-Process -Force 
}
catch {
    Write-Host "ImportData.xlsx currently not open on desktop." -ForegroundColor Red
}