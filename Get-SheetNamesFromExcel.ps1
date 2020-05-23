# ------------------------------------------------------------------------------
# Get-SheetNamesFromExcel
#   The function tells you what sheets that the Excel files have.
#   <PARAMETERS>
#     -File
#            System.IO.FileInfo object or a path string represents an Excel file.
#            You can use a pipeline to pass files to the function as below.
#            > ls -Filter "*.xlsx" | Get-SheetNamesFromExcel
#            Even if the input file is not an Excel format, the function opens 
#            it by Excel.Application and may return a sheet name or something.
#     -AskPassword
#            When you need type a password to open an Excel file, use the
#            parameter. You are asked to type a password just once.
#            The password is used for every protected file passed.
#     -FindLastCell
#            You can use the parameter to get the positions of the last used
#            cells in each sheet. This parameter adds two properties, LastRow and
#            LastColumn, onto the output object.
#     -FindLastCellIgnoringFormatted
#            The FindLastCellIgnoringFormatted finds the last cell with value in
#            contrast with that the FindLastCell finds the last cell formatted or
#            with value.
#   <OUTPUT>
#            PsCustomObject object that has two properties below.
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
        $File,
        [switch]
        $AskPassword,
        [switch]
        $FindLastCell,
        [switch]
        $FindLastCellIgnoringFormatted
    )
    begin{
        Set-Variable -Name "excel"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        Set-Variable -Name "password"
        $password = [System.Type]::Missing
        if($AskPassword){
            [securestring] $securedPassword = Read-Host -Prompt "Enter a password" -AsSecureString
            if($securedPassword.Length -gt 0) {
                $intptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedPassword)
                $password = [System.Runtime.InteropServices.Marshal]::PtrtoStringBSTR($intptr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($intptr)
            }
        }
    }
    process {
        Set-Variable -Name "book"
        # convert string to FileInfo when a string was passed
        if ($File -is [string]) {
            $File = [System.IO.FileInfo]::new($File)
            if(-not $File.Exists) {
                Write-Error "The file is not found. : $File"
                return
            }
        }
        # start processing a file
        Write-Verbose "processing a file : $($File.fullName)"
        if ($File -isnot [System.IO.FileInfo]) {
            Write-Verbose "The file is not a type of System.IO.FileInfo. : $($file.getType())"
            return
        }
        $missing = [System.Type]::Missing
        try {
            $book = $excel.Workbooks.Open(
                $file.FullName,        # FileName
                0,                     # UpdateLinks
                $true,                 # ReadOnly
                $missing,              # Format
                $password,             # Password
                $missing,              # WriteResPassword
                $true,                 # IgnoreReadOnlyRecommended
                $missing,              # Origin
                $missing,              # Delimiter
                $missing,              # Editable
                $missing,              # Notify
                $missing,              # Converter
                $missing               # AddToMru
            )
            # find the last cell used in each sheet
            return $book.sheets | ForEach-Object {
                Write-Verbose "    processing a sheet : $($_.Name)"
                $tmpObject = New-Object -TypeName PsObject -Property @{Sheet=$_.Name; File=$File}
                if($FindLastCell -or $FindLastCellIgnoringFormatted) {
                    # The point (x,y) is on the left-top in the UsedRange, which is formatted or with value
                    $y = $_.UsedRange.Row
                    $x = $_.UsedRange.Column
                    # The UsedRange
                    $rows = $_.UsedRange.Rows
                    $columns = $_.UsedRange.Columns
                    # Find the point (lastX,lastY)
                    Set-Variable -Name "lastY"
                    Set-Variable -Name "lastX"
                    # FindLastCell
                    if ($FindLastCell -and -not $FindLastCellIgnoringFormatted) {
                        $height = $rows.count
                        $width = $columns.count
                        $lastY = $y + $height - 1
                        $lastX = $x + $width - 1
                    }
                    # FindLastCellIgnoringFormatted
                    else {
                        # find the last row with value
                        for(
                            $height = $rows.count;
                            $height -gt 0;
                            $height--
                        ){
                            if($rows.Item($height).value2 -join "" -notlike ""){
                                break
                            }
                        }
                        if ($height -eq 0) {$lastY = 1} else {$lastY = $y + $height - 1}

                        # find the last column with value
                        for(
                            $width = $columns.count;
                            $width -gt 0;
                            $width--
                        ){
                            if($columns.Item($width).value2 -join "" -notlike ""){
                                break
                            }
                        }
                        if ($width -eq 0) {$lastX = 1} else {$lastX = $x + $width - 1}
                    }
                    Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name "LastRow" -Value $lastY
                    Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name "LastColumn" -Value $lastX
                }
                return $tmpObject
            }
        } catch {
            Write-Error "An error happened while reading the file"
            Write-Error $_.ErrorDetails
            Write-Error $_.ScriptStackTrace
            return
        } finally {
            if($book) {
                $book.Close()
            }
            Remove-Variable book
        }
    }
    end {
        # terminate an Excel process 
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null # return 0
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
# (c) to print the sheet names of an Excel file with password
#     > Get-SheetNamesFromExcel -File ./protected.xlsx -AskPassword
# (d) to get to know the used area roughly in each sheet of the Excel file
#     > Get-SheetNamesFromExcel -File ./protected.xlsx -AskPassword -FindLastCell
