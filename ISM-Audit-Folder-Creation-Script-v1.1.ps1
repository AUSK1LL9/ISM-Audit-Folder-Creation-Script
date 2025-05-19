.NOTES
Requires the Microsoft Excel COM object and the PSWriteWord module.
#>
#Required object for popup box:
Add-Type -AssemblyName Microsoft.VisualBasic

#Required to create docx files
#Install-Module PSWriteWord # Removed from here, best to do this manually

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Initialize variables for file paths and names
$excelFile = ""
$worksheetName = ""
$baseDirectory = ""

try {
    # Specify the path to the Excel file
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select ISM Excel File (E.g. C:\temp\RFFR SoA based on ISM December 2024.xlsx)"
    $openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
    $dialogResult = $openFileDialog.ShowDialog()

    # Check if the user clicked OK
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        # Get the selected file path
        $excelFile = $openFileDialog.FileName
    }
    else {
        Write-Warning "User cancelled file selection."
        exit  # Exit the script if the user cancels
    }

    # Open the Excel file
    $workbook = $excel.Workbooks.Open($excelFile)

    # Specify the sheet name
    $worksheetName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the ISM sheet name. If you aren't sure it will be in te format: ISM *Month* *Year* - ISM December 2024", "Enter ISM Sheet name")

    # Get the worksheet by name, with error handling
    try {
        $worksheet = $workbook.Sheets.Item($worksheetName)
    }
    catch {
        Write-Error "Worksheet '$worksheetName' not found in Excel file."
        exit # Exit if the worksheet is not found
    }

    # Specify the directory where new folders will be created
    $baseDirectory = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the path where you want the folder structure to be created. E.g. C:\Temp\ISM\", "Enter ISM folder creation directory")
    if (-not $baseDirectory)
    {
        Write-Warning "No directory provided. Exiting"
        exit
    }

    # Ensure the base directory ends with a backslash for consistency
    if (-not ($baseDirectory.EndsWith("\"))) {
        $baseDirectory += "\"
    }

    # Iterate through each row in the Excel sheet
    for ($row = 3; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
        # Extract values from columns A, B, D and E
        $columnA = $worksheet.Cells.Item($row, 1).Value2
        $columnB = $worksheet.Cells.Item($row, 2).Value2
        $columnC = $worksheet.Cells.Item($row, 3).Value2
        $columnD = $worksheet.Cells.Item($row, 4).Value2
        $columnE = $worksheet.Cells.Item($row, 5).Value2

        # Create the directory path based on values from columns B, and D
        $directoryPath = Join-Path -Path $baseDirectory -ChildPath "$columnA\$columnB\$columnC"

        # Create the directory if it doesn't exist
        if (!(Test-Path $directoryPath)) {
            try{
                New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
            }
            catch{
                 Write-Error "Failed to create directory: $directoryPath. Skipping row."
                 continue # Skip to the next row
            }

        }

        # Create a blank .docx file in each subfolder with a filename from column D
        $docxFileName = Join-Path -Path $directoryPath -ChildPath "$columnD.docx"

        # Check if the file already exists
        if (Test-Path $docxFileName) {
            Write-Warning "File '$docxFileName' already exists. Overwriting."
        }

        ### define new document
        try
        {
            $WordDocument = New-WordDocument -Path $docxFileName
        }
        catch
        {
            Write-Error "Failed to create Word document: $docxFileName. Skipping"
            continue
        }

        ### add heading and 1 sentence
        Add-WordText -WordDocument $WordDocument -Text $columnE -FontSize 18 -Bold $true
        Add-WordText -WordDocument $WordDocument -Text "" -FontSize 12
        Add-WordText -WordDocument $WordDocument -Text "* Add Evidence Here *" -FontSize 12

        ### Save document
        try
        {
             Save-WordDocument $WordDocument
        }
        catch
        {
            Write-Error "Failed to save document: $docxFileName"
        }
       #dispose
        $WordDocument.Dispose()

    }
    Write-Host "Script completed successfully."
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    # Close the Excel workbook and quit Excel, and release COM objects
    if ($workbook) {
        $workbook.Close([Type]::Missing, $true) #added to save changes.
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
   #garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
