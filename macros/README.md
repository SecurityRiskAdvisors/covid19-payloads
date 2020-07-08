# Macros

|Name|Description|
|--|--|
|Patchwork|Retrieves and executes a remote scriptlet file|
|Schutzwerk|Retrieves and executes a remote exe|
|Gamaredon|Performs local information collection, makes an HTTP request, then persists via a startup script|
|Talos|Drops and extracts an embedded zip file containing a Python interpreter and a Python script then executes the script using the interpreter|
|Kimsuky|Retrieves and executes a remote HTA file|
|Konni|Decodes and drops an embedded exe to disk then executes it|
|TA505|Decodes and drops an embedded xls to disk, extracts a DLL from the xls, then executes the dll using rundll32|

## Steps to Reproduce

### Patchwork

1. change URL in patchwork.scr to some arbitrary link (will be downloaded)
2. upload patchwork.scr file to some HTTP(S) site (e.g. gist.github.com)
3. create Excel doc and add URL from step 2 in cell A1
4. Save a document containing the patchwork.vba macro

### Schutzwerk

1. Compile http.cs
	- > cmd> c:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe -out:file.exe http.cs
2. Upload compiled binary to some HTTP(S) site
3. Put URL from step 2 in schutzwerk.vba
4. Save a document containing the schutzwerk.vba macro

### Gamaredon

1. Change URL in gamaredon.vba (anything benign will work - no data is being exfiltrated)
2. Save a document containing the gamaredon.vba macro as template (.dotm) and upload to some HTTP(S) site	
	- the macro uses MSXML to make HTTP requests -> in the macro editor, add a reference to "Microsoft XML, v6.0" (Tools -> References)
3. Set the URL to the document from step 2 as the template URL in a new document (Developer -> Document Template)

### Talos

1. Download Python embeddable (https://www.python.org/ftp/python/3.8.2/python-3.8.2-embed-amd64.zip)
2. Unzip the zip from step 1.
3. Make a new zip containing the "python-3.8.2-embed-amd64" folder and the provided talos.py file (both at the root of the archive)
	- note: if you use a more recent version of Python embeddable, this file name will be slighly different, make sure to modify the following line of the macro:
		- > Call Shell("""" & User & "\Python38\python-3.8.2-embed-amd64\python.exe" & """ """ & User & "\Python38\talos.py")
4. Get the file size (in bytes) of the zip and replace "7869792" with it in the following line of the macro:
	- > data = Right(data, 7869792)
5. Create a document (.doc/.docm) with the talos.vba macro
6. Concatenate the zip to the end of the document
	- cmd> copy /b file.doc+zip.zip newfile.doc

### Kimsuky

1. upload kimsuky.hta file to some HTTP(S) site
2. Put URL from step 1 in kimsuky.vba
3. Save a document containing the kimsuky.vba macro

### Konni

1. Compile http.cs
	- > cmd> c:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe -out:file.exe http.cs
2. Base64 encode the exe from step 1
3. Split the base64 encoded exe from step 2 into chunks appropriate for a macro 
	- can use something like: https://github.com/mdsecactivebreach/CACTUSTORCH/blob/master/splitvba.py
4. Replace the vbuffer in konni.vba with the chunked payload from step 3
	```
	'vbuffer = vbuffer + "foo"
	'should become
	
	vbuffer = vbuffer + "<chunk1>"
	vbuffer = vbuffer + "<chunk2>"
	```
5. Save a document (.doc/.docm) containing the edited konni.vba macro 
	- the macro uses MSXML to make HTTP requests -> in the macro editor, add a reference to "Microsoft XML, v6.0" (Tools -> References)

### TA505

1. Compile ta505.cpp
	- cmd> cl.exe /LD ta505.cpp
	- both ta505.cpp and ta505.h should be in the same folder
	- cl.exe is the Microsoft C++ compiler/linker; it can be downloaded from Microsoft here: https://visualstudio.microsoft.com/visual-cpp-build-tools/ 
	- this will generate ta505.dll
2. Create a new .xls file and embed the dll from step 1 into as an object (insert -> text dropdown -> object -> create from file)
3. Base64 the .xls from step 2 and perform same chunking as in Konni step 3 and step 4 but for the ta505.vba macro
4. Save an Execel (.xls/.xlsm) containing the edited ta505.vba macro


