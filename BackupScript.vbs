'Will assign the computer Name, Manufacturer, and Model to independent variables'

dim filesys, newfolder, newfolderpath, desktop, documents, pictures, downloads, instanceNumber

'Naming Folder Algorithum'

Dim arr(12), dt, x, y, Name, FolderName
Set wshShell = CreateObject( "WScript.Shell" )
Name = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
arr(1) = "Jan"        'Date'  
arr(2) = "Feb"        'Date' 
arr(3) = "Mar"        'Date'     
arr(4) = "Apr"        'Date' 
arr(5) = "May"        'Date'    
arr(6) = "Jun"        'Date' 
arr(7) = "Jul"        'Date' 
arr(8) = "Aug"        'Date'      
arr(9) = "Sep"        'Date' 
arr(10) = "Oct"       'Date'      
arr(11) = "Nov"       'Date'    
arr(12) = "Dec"       'Date' 

dt = now
x = month(dt)
y = year(dt)

FolderName = Name & " " & arr(x) & " " & y
instanceNumber = 0

'Sets the new folder name to be the Name of the computer'

newfolderpath = "c:\" & FolderName
rootpath = "c:\"
set filesys=CreateObject("Scripting.FileSystemObject")

If Not filesys.FolderExists(newfolderpath) Then
	Set newfolder = filesys.CreateFolder(newfolderpath)
	
Else 
	intAnswer = MsgBox(newfolderpath & " is already created." & "Do you want to delete the directory?", vbYesNo, "Delete Files")
	If intAnswer = vbYes Then
		instanceNumber = instanceNumber + 1
		filesys.DeleteFolder(newfolderpath)
		
	End If
End If
    
documents = "c:\Documents"
If Not filesys.FolderExists(documents) Then
	Set newfolder = filesys.CreateFolder(documents)
	
Else 
	intAnswer = MsgBox(documents & " is already created." & "Do you want to delete the directory?", vbYesNo, "Delete Files")
	If intAnswer = vbYes Then
		instanceNumber = instanceNumber + 1
		filesys.DeleteFolder(documents)
	End If
    
	
End If

pictures = "c:\Pictures"

If Not filesys.FolderExists(pictures) Then
	Set newfolder = filesys.CreateFolder(pictures)
	
Else 
	intAnswer = MsgBox(pictures & " is already created." & "Do you want to delete the directory?", vbYesNo, "Delete Files")
	If intAnswer = vbYes Then
		instanceNumber = instanceNumber + 1
		filesys.DeleteFolder(pictures)
	End If
	
End If

downloads = "c:\Downloads"

If Not filesys.FolderExists(downloads) Then
	Set newfolder = filesys.CreateFolder(downloads)
	
Else 
	intAnswer = MsgBox(downloads & " is already created." & "Do you want to delete the directory?", vbYesNo, "Delete Files")
	If intAnswer = vbYes Then
		instanceNumber = instanceNumber + 1
		filesys.DeleteFolder(downloads)
	End If

	
End If

desktop = "c:\Desktop"

If Not filesys.FolderExists(desktop) Then
	Set newfolder = filesys.CreateFolder(desktop)
	
Else 
	intAnswer = MsgBox( desktop & " is already created." & "Do you want to delete the directory?", vbYesNo, "Delete Files")
	If intAnswer = vbYes Then
		instanceNumber = instanceNumber + 1
		filesys.DeleteFolder(desktop)
	End If
	
End If

If instanceNumber > 0 Then
	Wscript.echo "Restarting script due to recent changes, please rerun the script and try again." & instanceNumber
	WScript.Quit 1
	
End If


'Logging Script'

Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
set f = fso.CreateTextFile("README", true)
fso.moveFile "README.txt", rootpath  'Throws a file not found error need to catch'
WScript.Echo f.created


'Creates a text file named README and writes info on the computer to the file'

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
  f.WriteLine "Name: " & objItem.Name
  f.WriteLine "Manufacturer: " & objItem.Manufacturer
  f.WriteLine "Model: " & objItem.Model
Next

'Creates a text file named README and writes info on the computer to the file'

'In README this will also provide the serial number to the file'

strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
For Each objSMBIOS in colSMBIOS
    f.WriteLine "Serial Number: " & objSMBIOS.SerialNumber
Next

'In README this will also provide the serial number to the file'

dt=now
f.WriteLine "Date of Backup: " & month(dt) & "-" &  day(dt) & "-" & year(dt)

'Copying Script'



'IF desktop.files.Count = 0 AND desktop.'


WScript.Echo "Script Complete"

