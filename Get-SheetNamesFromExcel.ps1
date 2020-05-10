﻿# ------------------------------------------------------------------------------
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
        $FindLastCell
    )
    begin{
        Set-Variable -Name "excel"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

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
                Write-Error "The file was not found. : $File"
                return
            }
        }
        # start processing a file
        Write-Verbose "processing a file : $($File.fullName)"
        if ($File -isnot [System.IO.FileInfo]) {
            Write-Verbose "The file is not a type of System.IO.FileInfo"
            return
        }
        $missing = [System.Type]::Missing
        try {
            $book = $excel.Workbooks.Open(
                $file.FullName,        # FileName
                0,                     # UpdateLinks
                $true,                 # ReadOnly
                $missing,              # Format
                $password ?? $missing, # Password
                $missing,              # WriteResPassword
                $true,                 # IgnoreReadOnlyRecommended
                $missing,              # Origin
                $missing,              # Delimiter
                $missing,              # Editable
                $missing,              # Notify
                $missing,              # Converter
                $missing               # AddToMru
            )
            # $sheetNames = $book.sheets | ForEach-Object {$_.name}
            # find the last cell used in each sheet
            return $book.sheets | ForEach-Object {
                Write-Verbose "    processing a sheet : $($_.Name)"
                $tmpObject = New-Object -TypeName PsObject -Property @{Sheet=$_.Name; File=$File}
                if($FindLastCell) {
                    $row = $_.UsedRange.Row + $_.UsedRange.Rows.count -1
                    $column = $_.UsedRange.Column + $_.UsedRange.Columns.count -1
                    Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name "LastRow" -Value $row
                    Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name "LastColumn" -Value $column
                }
                $tmpObject
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
