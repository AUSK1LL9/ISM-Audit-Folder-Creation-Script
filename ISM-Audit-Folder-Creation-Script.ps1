#required module to do better things wit docx easier
#https://evotec.xyz/hub/scripts/pswriteword-powershell-module/

#previous code had issue using word COM component - using PSWriteWord instead - Install-Module PSWriteWord
#Required object for popup box:
Add-Type -AssemblyName Microsoft.VisualBasic

#Required to create docx files
Install-Module PSWriteWord

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Specify the path to the Excel file
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Select ISM Excel File (E.g. C:\temp\RFFR SoA based on ISM December 2024.xlsx)"
$openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
$dialogResult = $openFileDialog.ShowDialog()
# Check if the user clicked OK
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
  # Get the selected file path
  $excelFile = $openFileDialog.FileName }

# Open the Excel file
$workbook = $excel.Workbooks.Open($excelFile)

# Specify the sheet name
$worksheet = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the ISM sheet name. If you aren't sure it will be in te format: ISM *Month* *Year* - ISM December 2024", "Enter ISM Sheet name")
# Get the worksheet by name
$worksheet = $workbook.Sheets.Item($worksheetName)

# Specify the directory where new folders will be created
$baseDirectory = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the path where you want thhe folder structure to be created. E.g. C:\Temp\ISM\", "Enter ISM folder creation directory")

# Iterate through each row in the Excel sheet
for ($row = 3; $row -le $worksheet.UsedRange.Rows.Count; $row++) {

    # Extract values from columns B, C, D and E
    $columnB = $worksheet.Cells.Item($row, 2).Value2
    # $columnC = $worksheet.Cells.Item($row, 3).Value2
    $columnD = $worksheet.Cells.Item($row, 4).Value2
    $columnE = $worksheet.Cells.Item($row, 5).Value2

    # Create the directory path based on values from columns B, C, and D
    $directoryPath = Join-Path -Path $baseDirectory -ChildPath "$columnB\ISM-$columnD"

    # Create the directory if it doesn't exist
    if (!(Test-Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }

    # Create a blank .docx file in each subfolder with a filename from column D
    $docxFileName = Join-Path -Path $directoryPath -ChildPath "ISM-$columnD.docx"

    ### define new document
    $WordDocument = New-WordDocument $docxFileName

    ### add heading and 1 sentance
    Add-WordText -WordDocument $WordDocument -Text $columnE -FontSize 21 -Bold $true
    Add-WordText -WordDocument $WordDocument -Text "" -FontSize 12
    Add-WordText -WordDocument $WordDocument -Text "Add Evidence *" -FontSize 12

    ### Save document
    Save-WordDocument $WordDocument
}

# Close the Excel workbook and quit Excel
$workbook.Close()
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Script completed successfully."
