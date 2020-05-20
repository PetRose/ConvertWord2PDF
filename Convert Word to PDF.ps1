# From this:
#     Source: https://stackoverflow.com/questions/46286292/powershell-word-to-pdf

# Adapted to Nets - MitID purposes by:
#          Initials     When           Why
#          PEROS        2020-03-16     To easier convert Word documents into PDF deliverables.
#                                      Observe, it does NOT require userintervention during conversion !
#          PEROS        2020-05-14     Set POSH up for running on tusta's Laptop (Torben Ulrich Stauer)
#                                      See also this PowerShell cheat sheet:
#                                          http://www.theochem.ru.nl/~pwormer/teachmat/PS_cheat_sheet.html
# Prerequisites:
#          A folder (and subfolders below) with Word Documents but WITHOUT any PDF files !
#          To avoid risky overwriting, the script is checking if there are PDF's in there
#          and if so: It will prompt and abandon the 'run'

#          Added Folder Select dialog, to ease.
#
# Information: It will take roughly 1 minute to convert 15 documents 
#  
Add-Type -AssemblyName System
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$InitialDirectory = $Env:OneDriveCommercial     # Could be candidate for InitialFolder (RootFolder parameter)


$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog   
#    InitialDirectory = $InitialDirectory 

$null = $FileBrowser.ShowDialog()               # Kick in to select dialogue

$desiredFolder = $FileBrowser.Selectedpath      # Pikc up the chosen one

 
New-Variable pdfno -Value 0 -Description "Holds the number of PDF files in the input folder, if any is found in there. Once converted, we pick up number once more."
$path = $desiredFolder                          # InitialDirectory  

Write-Host "Checking prerequisites for PDF conversion.." -ForegroundColor Red -BackgroundColor White   

#Get-ChildItem -Path $path -Filter *.PDF -Recurse | Set-Variable pdfno
Get-ChildItem -Path $path -Filter *.PDF -Recurse | Measure-Object | select-Object -Property Count | Set-Variable pdfno
If ($pdfno.count -gt 0) 
    {
    Write-Host "You still have "$pdfno.Count" PDF files in the target folder:" -ForegroundColor Red -BackgroundColor White
    Write-Host "   "$path -ForegroundColor Red -BackgroundColor White
    Write-Host "Because of this, we have to exit early for this conversion. Remove the PDF files from folders, and retry." -ForegroundColor Red -BackgroundColor White
    
    Remove-Variable -Name pdfno    # Cleanup
    Break                          # Stop immediately
    }

Write-Host "Starting PDF Conversion of folder:" -ForegroundColor Red -BackgroundColor White
Write-Host "   "$path -ForegroundColor Red -BackgroundColor White
Get-Date | Write-Host -ForegroundColor Red -BackgroundColor White


$wd = New-Object -ComObject Word.Application
Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse |
    ForEach-Object {
        $doc = $wd.Documents.Open($_.Fullname)
        $pdf = $_.FullName -replace $_.Extension, '.pdf'
        $doc.ExportAsFixedFormat($pdf,17,$false,0,3,1,1,0,$false, $false,0,$false, $true)
        $doc.Close(0)                  # Close without Saving, can perhaps avoid the dreaded ReadOnlyRecommended
    }
$wd.Quit()

Get-ChildItem -Path $path -Filter *.PDF -Recurse | Measure-Object | select-Object -Property Count | Set-Variable pdfno

Write-Host "Ending PDF Conversion at:" -ForegroundColor Red -BackgroundColor White
Get-Date | Write-Host -ForegroundColor Red -BackgroundColor White
Write-Host "Number of Word documents converted to PDF : "$pdfno.Count -ForegroundColor Red -BackgroundColor White

Remove-Variable -Name pdfno
Start-Sleep -s 15      # To allow the user to see the Results of Conversion, before ending...
Return 