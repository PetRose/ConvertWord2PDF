<#
.SYNOPSIS
    This CMDlet converts a designated folder and its sub-folders with Word documents to PDF documents in the same folder(s).

 
.DESCRIPTION
    Without the optional parameter, the CMDlet will prompt user for the InputFolder to use.
    If InputFolder parameter is given, it will attempt to process Word documents in that folder and its sub-folders (recurse is default)

    To avoid risky overwriting, the script is checking if there are any PDF's in the designated folder(s)
    and if so: It will prompt and abandon the 'run'
    If InputFolder parameter is given and the folder does not exist it will abandon run with a message.
    
    Otherwise, it will process the Word documents if any and save them as PDF. 
    Messages goes to standard output. 

    If message output is redirected (see EXAMPLES) and a valid InputFolder is given, you can run the CMDlet as a Scheduled Task
    that will convert all Word files to PDF in the designated folder(s).
    When this is set up as a scheduled task, one has to Move the created PDF files away from the input Folder(s) before triggering the same task again.
    This could be made with the simple Move-Item PS commandlet. See Help for this, on how this can be done

.PARAMETER InputFolder
    InputFolder can be specified if you want to avoid the dialog session to help pointing out a folder with Word documents to use for conversion.
    In this way, the script could be run without user intervention.

.INPUTS
    A folder and its sub-folders (a kind of recurse is defaulted) containg Word documents to be converted.
    The folder(s) must not contain any PDF files prior to calling the CMDlet.

.OUTPUTS
    CMDlet returns a set of PDF files one for each found Word document that has been converted, in the same folder(s) processed.
    
    ReturnCode 0  : No documents were converted to PDF. There might not be any Word documents, or some odd error occurred. See Error Output
    ReturnCode 1  : One or more documents are converted to PDF documents.
    ReturnCode 4  : User abandoned the process by Cancelling the Folder Selction dialog.
    ReturnCode 8  : The InputFolder to convert, does not exist.
    ReturnCode 12 : There is still PDF documents in the designated InputFolder or its sub-folders. Script stops.

.EXAMPLE
    Convert-Word2PDF.ps1  
    Without parameter, it will open a Dialog window that allow the user to select the folder for which the conversion have to work with. 
    Remember, sub-folders are processed too.

.EXAMPLE
    Convert-Word2PDF.ps1 -InputFolder 'C:\tmp\all'
    Since the InputFolder parameter is given, it will check if the folder exist. If the folder does not exist it will abandon run with a message.
    If there exist PDF files inside the folder or its sub-folders, it will abandon too.
    If no PDF files exists in folder(s) it will process the Word documents if any, and save them as PDF.

.EXAMPLE
    A PS commandline version:
      .\Convert-Word2PDF.ps1 -inputfolder 'D:\Temp\Word-Testing-Convert2PDF' *> script.log
    The script will not prompt for Folder as its given by the parameter. 
    Also, there will be no message output displayed, instead messages (all types) will go to the file 'script.log' inside folder where the script is run.


.NOTES
    NAME:       Convert-Word2PDF
    AUTHOR:     Peter Rosenberg, peter.rosenberg.dk@gmail.com
    CREATEDATE: May, 2020
    UPDATE:     May 26, 2020 by Peter Rosenberg
                To make it possible to use InputFolder as parameter and then allow a truely Batch operation possible.
    THROUGHPUT: It will take roughly 1 minute to convert 15 medium sized documents 

    CREDITS
                From this, snippets of the solution is found:
                  Source: https://stackoverflow.com/questions/46286292/powershell-word-to-pdf
.LINK
    GitHub repository: https://github.com/PetRose/ConvertWord2PDF

#>

#Requires -Version 5.1

[CmdletBinding()]

Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)][string]$InputFolder
    ) 

# ------------------------- [Cmdlet start] -----------------------------
#  
Add-Type -AssemblyName System
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$InitialDirectory = $Env:OneDriveCommercial     # Could be candidate for InitialFolder (RootFolder parameter, if using OneDrive for Business/School)

If ($InputFolder -eq '') 
    {
    $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog   
#    InitialDirectory = $InitialDirectory 

    $null = $FileBrowser.ShowDialog()               # Kick in to select dialogue

    $desiredFolder = $FileBrowser.Selectedpath      # Pikc up the chosen one

    If ($FileBrowser.Selectedpath -eq '') {
        Write-Host "You cancelled the operation. Script stops here." -ForegroundColor Red -BackgroundColor White  
        Return 4                         # Stop immediately
        }
    } else {
            if ( (Test-Path -Path $inputfolder -PathType Container) -eq $false) {
                 Write-Host "Folder given as parameter:"$InputFolder" does not exist. Script stops here." -ForegroundColor Red -BackgroundColor White  
                 Return 8                # Stop immediately
                 }     
            # Continue with a folder that exists
            $DesiredFolder= $inputFolder
        }
 
New-Variable pdfno -Value 0 -Description "Holds the number of PDF files in the input folder, if any is found in there. Once converted, we pick up number once more."
$path = $desiredFolder                         

Write-Host "Checking prerequisites for PDF conversion.." -ForegroundColor Red -BackgroundColor White   

Get-ChildItem -Path $path -Filter *.PDF -Recurse | Measure-Object | select-Object -Property Count | Set-Variable pdfno
If ($pdfno.count -gt 0) 
    {
    Write-Host "You still have "$pdfno.Count" PDF files in the target folder:" -ForegroundColor Red -BackgroundColor White
    Write-Host "   "$path -ForegroundColor Red -BackgroundColor White
    Write-Host "Because of this, we have to exit early for this conversion. Remove the PDF files from folders, and retry." -ForegroundColor Red -BackgroundColor White
    
    Remove-Variable -Name pdfno    # Cleanup
    Return 12                      # Stop immediately
    }

Write-Host "Starting PDF Conversion of folder:" -ForegroundColor Red -BackgroundColor White
Write-Host "   "$path -ForegroundColor Red -BackgroundColor White
Get-Date | Write-Host -ForegroundColor Red -BackgroundColor White


$wd = New-Object -ComObject Word.Application
Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse |
    ForEach-Object {
        $doc = $wd.Documents.Open($_.Fullname,$false,$true)     # See if we can open it ReadOnly (for .doc Word files) 
        Write-Host "   Converting: "$_.Fullname -ForegroundColor Red -BackgroundColor White
        $pdf = -join @($path,[System.IO.Path]::AltDirectorySeparatorChar,($_.Name -replace $_.Extension, '.pdf'))   # Bugfix 2020-05-20
        $doc.ExportAsFixedFormat($pdf,17,$false,0,3,1,1,0,$false, $false,0,$false, $true)
        $doc.Close(0)                  # Close without Saving, can perhaps avoid the dreaded ReadOnlyRecommended
    }
$wd.Quit()

Get-ChildItem -Path $path -Filter *.PDF -Recurse | Measure-Object | select-Object -Property Count | Set-Variable pdfno

Write-Host "Ending PDF Conversion at:" -ForegroundColor Red -BackgroundColor White
Get-Date | Write-Host -ForegroundColor Red -BackgroundColor White
Write-Host "Number of Word documents converted to PDF : "$pdfno.Count -ForegroundColor Red -BackgroundColor White

Start-Sleep -s 10      # To allow the user to see the Results of Conversion, before ending...

# Setting a return code that can be used, if someone would do countermeasures when running as background batch.

if ((get-Variable pdfno -ValueOnly).count -eq 0) {Return 0}    # No documents converted. There might not be any Word documents at all.
   else {Return 1}                                             # There are at least one document converted. These has to be moved away before next conversion is attempted