# ConvertWord2PDF

**Description**

The purpose of this handy tool, is to facilitate the conversion process for large scale projects/organisations which:
- Produce large amount of Word documents, and has to mass-convert these for audiences in the form af PDF files readable through
the Adobe Reader or any PDF Reader.

*Prerequisites*

A PowerShell exucution platform, usually available in Windows 8 or later versions like Windows 10.
If you are unaware of what you have, please make these checks:
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

        -- $env:UserProfile"\Documents\WindowsPowerShell\Modules" as the folder where you can place you Modules if you are into coding of these.
        -- $env:UserProfile"\Documents\WindowsPowerShell" as folder for your scripts or any sub-folder for that matter

    So the last thing to check, is to see if you can execute PowerShell scripts. Try doing this command:

            $env:PSExecutionPolicyPreference
        
        If output says 'RemoteSigned' you are good to go.
        
        Otherwise do this command:
        
            Set-ExecutionPolicy -ExecutionPolicy -Scope CurrentUser

**Download the Conversion Script**

You can simply download it as Script, bu beware yoyr browser may tell you have to accept a prompt telling you it could be dangerous.

You can also just Copy the code inside the Script, and you can paste it into a file to name with filetype '.PS1' and place it in the designated Script folder you decide.

**Activating the Convert 2 PDF Script**

    Work in progress

