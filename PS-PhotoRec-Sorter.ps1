<#
.SYNOPSIS
    This script will sort files recovered by Photorec into folders by file extension based on Photorec's 'report.xml' and 
	generate an Excel report.
.DESCRIPTION
    An Excel version of Photorec's report.xml will be created and recovered files will be sorted into folders by extension.
    Files which cannot be located or not listed in report.xml will be left as-is.
    Place and run script from 'recup_dir' parent directory.
.EXAMPLE
    C:\PS> Generate-Photorec-Excel-Report.ps1
    Sorts Photorec recovered files into folders by file extension and generates Excel version of report.xml.
.NOTES
    Author: Y.G.
	GitHub: @varbytes
    Date:   6 May 2025
#>


# Find Photorec report.xml file
$curPath = (Get-Location).Path
$xmlFilePath = Join-Path -Path $curPath -ChildPath "recup_dir.1\report.xml"
if (!(Test-Path $xmlFilePath)) {
    Write-Host "[ERROR] '${xmlFilePath}' does not exist!"
    Exit 1
}

Try {

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    Start-Transcript -IncludeInvocationHeader -Path "Photorec-Excel-Report_$timestamp.log"

    # Create Excel COM Object and worksheet
    $excelFile = New-Object -ComObject Excel.Application
    # change $False to $True during debugging to see Excel file
    $excelFile.Visible = $False
    $excelWorkbook = $excelFile.Workbooks.Add(1)
    $excelWorksheet = $excelWorkbook.Worksheets.Item(1)
    $excelWorksheet.Name = "Report"

    Write-Host "Reading contents of '$xmlFilePath'"
    [xml]$xmlFile = Get-Content $xmlFilePath
    Write-Host

    # Setup header row
    $excelWorksheet.Cells.Item(1, 1) = "S/N"
    $excelWorksheet.Cells.Item(1, 2) = "Filename"
	$excelWorksheet.Cells.Item(1, 3) = "PhotoRec Filename"
    $excelWorksheet.Cells.Item(1, 4) = "File Ext"
    $excelWorksheet.Cells.Item(1, 5) = "File Size"
    $excelWorksheet.Cells.Item(1, 6) = "MD5 Hash"
    $excelWorksheet.Cells.Item(1, 7) = "Byte Runs"

    # Loop through each fileobject in Photorec results
    $x = 2
    foreach ($entry in $xmlFile.dfxml.fileobject) {
        # Locate carved file
        $filepath = $entry.filename
		if ($filepath.contains("recup_dir")) {
            $filepath = $filepath.split("/",3)[2]
        } else {
            $filepath = "recup_dir.1/" + $filepath
        }
        if (!(Test-Path $filepath)) {
            Write-Host "[ERROR] Could not find file '${filepath}'. Trying closest match..."
			$tempname = $filepath.Substring(0, $filepath.lastIndexOf("."))
			$tempext = $filepath.split(".")[-1]
			$filepath = -join($tempname, "_*.", $tempext)
			if (!(Test-Path $filepath)) {
				Write-Host "[ERROR] No match found for '${filepath}'."
				Continue
			}
        }

        # Get carved file and move file into folder
        $file = Get-Item -Path $filepath
		$fileExt = $file.Extension
		Write-Host "[INFO] Hashing '$($file.Name)'."
        $fileHash = (Get-FileHash -Path $file.FullName -Algorithm MD5).Hash.ToLower()
        $dstPath = Join-Path -Path $curPath -ChildPath $fileExt.ToUpper()
        if (!(Test-Path $dstPath -PathType Container))
	    {
		    New-Item -ItemType Directory -Path $dstPath
	    }
        Move-Item -Path $file.FullName -Destination $dstPath

        # Print details of carved file to Excel sheet
        $excelWorksheet.Cells.Item($x, 1) = $x - 1
        $excelWorksheet.Cells.Item($x, 2) = $file.Name
		$excelWorksheet.Cells.Item($x, 3) = $entry.filename
        $excelWorksheet.Cells.Item($x, 4) = $fileExt
        $excelWorksheet.Cells.Item($x, 5) = $entry.filesize
        $excelWorksheet.Cells.Item($x, 6) = $fileHash
        $byterunstr = ""
        foreach ($byterun in $entry.byte_runs.byte_run) {
            $byterunstr += "offset='" + $byterun.GetAttribute("offset") + "' "
            $byterunstr += "img_offset='" + $byterun.GetAttribute("img_offset") + "' "
            $byterunstr += "len='" + $byterun.GetAttribute("len") + "'`n"
        }
        # trim trailing semi-colon and newline
        $excelWorksheet.Cells.Item($x, 7) = $byterunstr.TrimEnd(";","`n")
        $x++
    }

    # Beautify Excel file with header row formatting and column widths
    $range = $excelWorksheet.Range("A1:G1")
    $range.Font.Bold = $True
    $excelWorksheet.Cells.Item(1,1).ColumnWidth = 8
    $excelWorksheet.Cells.Item(1,2).ColumnWidth = 40
	$excelWorksheet.Cells.Item(1,3).ColumnWidth = 40
    $excelWorksheet.Cells.Item(1,4).ColumnWidth = 10
    $excelWorksheet.Cells.Item(1,5).ColumnWidth = 15
    $excelWorksheet.Cells.Item(1,6).ColumnWidth = 40
    $excelWorksheet.Cells.Item(1,7).ColumnWidth = 60

    # Save the Excel object as file
    $excelFile.DisplayAlerts = $False
    $excelWorkbook.SaveAs((Join-Path -Path $curPath -ChildPath "photorec-results.xlsx"))
    $excelWorkbook.Close($False)
    $excelFile.Quit()

    # Move the Photorec XML report out to base folder
    Move-Item -Path $xmlFilePath -Destination $curPath

    Write-Host
    Write-Host "##### Sorted $(${x}-2) carved files from report.xml"
    Write-Host
    Stop-Transcript
    Write-Host
    Write-Host "'photorec-results.xlsx' and 'report.xml' can be found in '$curPath'"

}
Catch
{
    Write-Host("TRAPPED: " + $_.ItemName)
    Write-Host("[ERROR!] " + $_.Exception.Message)
    continue  # continue onto the "Finally" block
}
Finally
{
    # Close the COM object
	$excelFile.Quit()
}