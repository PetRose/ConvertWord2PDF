# ConvertWord2PDF

## Description

The purpose of this handy tool, is to facilitate the conversion process for large scale projects/organisations which:
- Produce large amount of Word documents, and has to mass-convert these for audiences in the form af PDF files readable through
the Adobe Reader or any PDF Reader.

### Prerequisites

A PowerShell exucution platform, usually available in Windows 8 or later versions like Windows 10.
You also need Word Application, I have developed this for Word 10 and also tested it with Word2016. 
Be aware that some Word10 'features' makes the .doc documents readonly which could interfear with the ability  
to convert it, but lets hope for the best. 

If you are unaware of what you have for PowerShell, please make these checks:
 - Open PowerShel ISE (from Start menu)
 - In the Commandline near bottom, type this and press Enter: 
        $PSVersionTable

    Output could look like this:

        Name                           Value
        ----                           -----
        PSVersion                      5.1.18362.752
        PSEdition                      Desktop
        PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}
        BuildVersion                   10.0.18362.752
        CLRVersion                     4.0.30319.42000
        WSManStackVersion              3.0
        PSRemotingProtocolVersion      2.3
        SerializationVersion           1.1.0.1
- The first line (PSVersion) is the most important. In above example, it shows version is 5.1 at least.
If you do not have PowerShell at all. please google how to get it.

- Determine where your PowerShell scripts could be placed. Do this command in the same commamdline as above:

    ls  $env:UserProfile"\Documents\WindowsPowerShell" 

    The output could look like this:

        Directory: C:\Users\Peter\Documents\WindowsPowerShell

        Mode                LastWriteTime         Length Name
        ----                -------------         ------ ----
        d-----       17/05/2020     16.36                Modules
        d-----       20/05/2020     10.43                Scripts
        -a----       16/01/2020     17.07           1545 Microsoft.PowerShell_profile.ps1
- And the follow the command by this: 

    Get-ChildItem Env:PSModulePath | FT -wrap
  
    Output could look loke this:
  
    Name                           Value
    ----                           -----
    PSModulePath                   C:\Users\Peter\Documents\WindowsPowerShell\Modules;C:\Program
                                   Files\WindowsPowerShell\Modules;C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\;C:\opscode\chefdk\modules\;C:\Program Files (x86)\Microsoft
                                    SDKs\Azure\PowerShell\ServiceManagement;c:\Users\Peter\.vscode\extensions\ms-vscode.powershell-2020.4.0\modules
    
    Above two commands shows, that you can use:

       $env:UserProfile"\Documents\WindowsPowerShell\Modules" as the folder where you can place you Modules if you are into coding of these.
       $env:UserProfile"\Documents\WindowsPowerShell" as folder for your scripts or any sub-folder for that matter

The last thing to check, is to see if you can execute PowerShell scripts. Try doing this command:

            $env:PSExecutionPolicyPreference
        
If output says 'RemoteSigned' you are good to go.
Otherwise do this command:
        
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

## Download the Conversion Script
You can simply download it as Script, bu beware your browser may tell you have to accept a prompt telling you it could be dangerous.
Once this is done, you have to "Unblock" the file to tell PowerShell its OK to execute on your system:

    You right-click on the file to get its properties, and you can check the 'Unblock' checkbox near the bottom on first tab.

You can also just Copy the code inside the Script, and then paste it into a text file you save with a name with filetype '.PS1' and place it in the designated Script folder you decide. I suggest you name it *Convert2PDF.PS1*

## Activating the Convert 2 PDF Script
You can do this from either the PowerShell ISE, pointing to or opening the Script and chose the 'Run' or even change directory inside the PS Commandline 
to your script folder, and then execute it from there.

Or for the end user, you can create a shortcut to the Script and just let the Operating system to start PowerShell and let it execute the Script.

The first thing happening, when script starts, it lets you decide which folder(s) to work with this Conversion.

Once you have decided a folder and press OK, then you may see this kind of message:

        Checking prerequisites for PDF conversion..
        You still have  87  PDF files in the target folder:
            Because of this, we have to exit early for this conversion. Remove the PDF files from folders, and retry.

This is a precaution, so you do not convert and possibly overwrite other PDF files inadvertently.

In other words, #you need to have folders without PDF files* in order to process the Word files in them.

Also, notice the ***sub-folders of the folder you chose, with Word files will be converted !***

A succesful conversion would show up like this (small batch):
      Checking prerequisites for PDF conversion..
      Starting PDF Conversion of folder:
         C:\Users\SonjaPC\Documents\TestWordConv\Word
      24-05-2020 18:36:53
         Converting:  C:\Users\SonjaPC\Documents\TestWordConv\Word\Aneopdeling.doc
         Converting:  C:\Users\SonjaPC\Documents\TestWordConv\Word\Organisering af fotos.docx 
      Ending PDF Conversion at:
      24-05-2020 18:36:55
      Number of Word documents converted to PDF :  2

So *happy Conversion* with your word and documentation generation efforts

Cheers
Peter Rosenberg
Author, and owner of Rosenberg-IT


