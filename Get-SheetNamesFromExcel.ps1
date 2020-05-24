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
#     -GetCellByPosition
#            Use the parameter to get values of the cells you know its location.
#            You specify the names and positions of the cells like below.
#            > -GetCellByPosition @{ScreenTitle="B1";Author="J1"}
#            The parameter adds properties whose name is the specified name and
#            whose value is the value of the cell at the specified position.
#     -GetCellByLeftTitle
#            Use the parameter to get values of the cells you know the label
#            in its left. Specify the names and labels like below.
#            > -GetCellByLeftTitle @{ScreenTitle="SCREEN NAME";Author="BY"}
#            The parameter adds properties whose name is the specified name and
#            whose value is the value of the cell reached by pressing keys
#            [End]+[Right] at the cell with the title.
#     -GetCellByTopTitle
#            Use the parameter to get values of the cells you know the label
#            in its above. Specify the names and labels like below.
#            > -GetCellByTopTitle @{ScreenTitle="SCREEN NAME";Author="BY"}
#            The parameter adds properties whose name is the specified name and
#            whose value is the value of the cell reached by pressing keys
#            [End]+[Down] at the cell with the title.
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
        $FindLastCellIgnoringFormatted,
        [Hashtable]
        $GetCellByPosition,
        [Hashtable]
        $GetCellByLeftTitle,
        [Hashtable]
        $GetCellByTopTitle
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
                # Get the values of the cells in the specified positions
                if ($null -ne $GetCellByPosition) {
                    foreach ($key in $GetCellByPosition.Keys) {
                        $value = $_.Range($GetCellByPosition[$key]).Value2
                        Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name $key -Value $value
                        Write-Debug "        GetCellByPosition: KEY=$key VALUE=$($GetCellByPosition[$key])) CELL=$($value)"
                        Clear-Variable -Name "value"
                    }
                }
                # Get the values of the cells which come with the specified title in its left
                if ($null -ne $GetCellByLeftTitle) {
                    foreach ($key in $GetCellByLeftTitle.Keys) {
                        $title = $_.Cells.Find($GetCellByLeftTitle[$key])
                        Set-Variable -Name "value"
                        if ($null -ne $title) {
                            # can not find the type [Microsoft.Office.Interop.Excel.XlDirection]::XlToRight
                            $value = $title.End(-4161).Value2
                            Write-Debug "        GetCellByLeftTitle: KEY=$key VALUE=$($GetCellByLeftTitle[$key])) CELL=$($value)"
                        }
                        Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name $key -Value $value
                        Clear-Variable -Name "value"
                    }
                }
                # Get the values of the cells which come with the specified title in its above
                if ($null -ne $GetCellByTopTitle) {
                    foreach ($key in $GetCellByTopTitle.Keys) {
                        $title = $_.Cells.Find($GetCellByTopTitle[$key])
                        Set-Variable -Name "value"
                        if ($null -ne $title) {
                            # can not find the type [Microsoft.Office.Interop.Excel.XlDirection]::XlDown
                            $value = $title.End(-4121).Value2
                            Write-Debug "        GetCellByTopTitle: KEY=$key VALUE=$($GetCellByTopTitle[$key])) CELL=$($value)"
                        }
                        Add-Member -InputObject $tmpObject -MemberType NoteProperty -Name $key -Value $value
                        Clear-Variable -Name "value"
                    }
                }
                return $tmpObject
            }
        } catch {
            Write-Error "An error happened while reading the file"
            Write-Error $_.ErrorDetails
            Write-Error $_.ScriptStackTrace
            Write-Error $_.InvocationInfo
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
#     > Import-Module Get-SheetNamesFromExcel.ps1 -Force
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
