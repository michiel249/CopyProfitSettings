'* Michiel Singor
'* 2015

dim filesys
Const OverWriteFiles = True
set filesys = CreateObject("Scripting.FileSystemObject") 

Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" ) 
'* Fill variable strUserName with %username%

strComputername = Inputbox("Please proivde a computername to copy settings from:","Proift settings copy")
'* Fill strComputername with user provided remote computername

strPROFITPR = "\\" & strComputername & "\C$\Users\" & strUserName & "\AppData\Roaming\AFAS Software"
strPROFITPL = "C:\Users\" & strUserName & "\AppData\Roaming\AFAS Software"
'* Fill variable strPROFITPL and strPROFITPR
'* strPROFITPR = profit remote path \\computername\c:\users\%userprofile%\appdata\....
'* strPROFITPL = profit local path C:\%userprofile%\appdata\...

'* WScript.Echo strPROFITPR
'* WScript.Echo strPROFITPL

If filesys.FolderExists(strPROFITPR) Then
filesys.CopyFolder strPROFITPR, strPROFITPL, OverWriteFiles
MsgBox "Profit settings copy successful!" , vbInformation , "Copy successful"
Else
MsgBox "Oops.. something went wrong" & vbCrLf & "is this correct? ? Computername: " & strComputername & " Username: " & strUserName, vbCritical , "Oops.. something went wrong"
End If
