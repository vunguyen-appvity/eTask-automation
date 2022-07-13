$dataExcel = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName Update
$dataExcelCompare = Import-Excel -Path "C:\eTaskAutomationTesting\ImportData.xlsx" -WorksheetName 'Update (2)'

$b = Compare-Object -ReferenceObject $dataExcel -DifferenceObject $dataExcelCompare -Property status -IncludeEqual -PassThru
foreach ($r in $b) {
    if ($r.SideIndicator -eq '<=') {
        $r
    }
    elseif ($r.SideIndicator -eq '=>') {   
    }
}
