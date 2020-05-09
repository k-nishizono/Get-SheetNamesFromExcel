# ------------------------------------------------------------------------------
# Get-SheetNamesFromExcel
#   The function tells you what sheets that the Excel files have.
#   input  : System.IO.FileInfo object that represents an Excel file.
#            You can use a pipeline to pass files to the function as below.
#            > ls -Filter "*.xlsx" | Get-SheetNamesFromExcel
#            Even if the input file is not an Excel format, the function opens 
#            it by Excel.Application and may return a sheet name or something.
#   output : PsObject object that has two properties below.
#             - "Sheet" : the name of a sheet in the Excel file
#             - "File"  : System.IO.FileInfo object that were input
#
# Author : Zono
# ------------------------------------------------------------------------------

Set-StrictMode -Version latest

function global:Get-SheetNamesFromExcel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $file
    )
    begin{
        Set-Variable -Name "excel"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    }
    process {
        Set-Variable -Name "book"
        Set-Variable -Name "sheetNames"
        Write-Verbose "processing a file : $($file.fullName)"
        if ($file -isnot [System.IO.FileInfo]) {
            Write-Verbose "The file is not a type of System.IO.FileInfo"
            return
        }
        $missing = [System.Type]::Missing
        try {
            $book = $excel.Workbooks.Open(
                $file.FullName, # FileName
                0,              # UpdateLinks
                $true,          # ReadOnly
                $missing,       # Format
                $missing,       # Password
                $missing,       # WriteResPassword
                $true,          # IgnoreReadOnlyRecommended
                $missing,       # Origin
                $missing,       # Delimiter
                $missing,       # Editable
                $missing,       # Notify
                $missing,       # Converter
                $missing        # AddToMru
            )
            $sheetNames = $book.sheets | %{$_.name}
            $book.Close()
        } catch {
            if($book){ 
                $book.Close()
                Remove-Variable book
            }
            return
        }
        remove-variable book
        Write-Verbose "Sheet Names : $sheetNames"
        # Add-Member -InputObject $file -MemberType NoteProperty -Name "sheetNames" -Value $sheetNames
        $sheets = $sheetNames | %{New-Object -TypeName PsObject -Property @{Sheet=$_; File=$file}}
        Remove-Variable sheetNames
        return $sheets
    }
    end {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Remove-Variable excel
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# INSTALLATION
# (i) to define the function Get-SheetNamesFromExcel to call this file
#     > ./Get-SheetNamesFromExcel.ps1
# USAGE EXAMPLE
# (a) to print the sheet names of the Excel files on the current location
#     > ls *.xlsx | Get-SheetNamesFromExcel | %{$_.Sheet}
# (b) to search Excel files recursively on the current location 
#     and write the sheet names and file paths onto a csv file hogehoge.csv
#     > ls -Recurse -Filter "*.xlsx" | Get-SheetNamesFromExcel | export-csv -path ./hogehoge.csv
