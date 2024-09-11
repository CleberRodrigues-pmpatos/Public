Option Explicit
Const HOST_DEFAULT = "cscript.exe"

Private ShellApp
Private oShell
Private FileSystemObject
Private WinMgmtsCIMV2
Private XMLHTTP
Private AttemptedCreateXMLHTTP
Private FileStream

Set ShellApp = Nothing
Set oShell = Nothing
Set FileSystemObject = Nothing
Set WinMgmtsCIMV2 = Nothing
Set XMLHTTP = Nothing
Set AttemptedCreateXMLHTTP = Nothing
Set FileStream = Nothing
AttemptedCreateXMLHTTP = False

Function GetFileExtensionFromContentType(contentType)
	Select Case LCase(contentType)
		Case "application/octet-stream"
			GetFileExtensionFromContentType = "exe"
		Case "application/x-msi"
			GetFileExtensionFromContentType = "msi"
		Case "application/pdf"
			GetFileExtensionFromContentType = "pdf"
		Case "image/jpeg"
			GetFileExtensionFromContentType = "jpg"
		Case "image/png"
			GetFileExtensionFromContentType = "png"
		Case "text/html"
			GetFileExtensionFromContentType = "html"
		Case "application/zip"
			GetFileExtensionFromContentType = "zip"
		Case Else
			GetFileExtensionFromContentType = ""
	End Select
End Function

Function GetFileStream()
	If FileStream Is Nothing Then Set FileStream = CreateObject("ADODB.Stream")
	Set GetFileStream = FileStream
End Function

Function GetFileSystemObject()
	If FileSystemObject Is Nothing Then Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
	Set GetFileSystemObject = FileSystemObject
End Function

Function GetShell()
    If oShell Is Nothing Then Set oShell = CreateObject("WScript.Shell")
    Set GetShell = oShell
End Function

Function GetShellApp()
	If ShellApp Is Nothing Then Set ShellApp = CreateObject("Shell.Application")
	Set GetShellApp = ShellApp
End Function

Function GetXMLHTTP()
	If XMLHTTP Is Nothing And AttemptedCreateXMLHTTP = False Then
		Dim versions, objHTTP, i
		versions = Array("WinHttp.WinHttpRequest.5.1", "MSXML2.XMLHTTP.6.0", "MSXML2.ServerXMLHTTP.6.0", "MSXML2.XMLHTTP.3.0", "MSXML2.ServerXMLHTTP.3.0", "MSXML2.XMLHTTP", "Microsoft.XMLHTTP")

		Set objHTTP = Nothing

		For i = 0 To UBound(versions)
			On Error Resume Next
			Set objHTTP = CreateObject(versions(i))
			If Err.Number = 0 Then
				Set XMLHTTP = objHTTP
				Exit For
			End If
		Next
		AttemptedCreateXMLHTTP = True
	End If
	Set GetXMLHTTP = XMLHTTP
End Function

Function GetWinMgmtsCIMV2()
	If WinMgmtsCIMV2 Is Nothing Then Set WinMgmtsCIMV2 = GetObject("WinMgmts:\\.\root\CIMV2")
	Set GetWinMgmtsCIMV2 = WinMgmtsCIMV2
End Function

Function GetArguments()
	Dim strArguments, i
	strArguments = ""
	
	For i = 0 To WScript.Arguments.Count - 1
		strArguments = strArguments & " """ & WScript.Arguments(i) & """"
	Next

	GetArguments = strArguments
End Function

Function GetOSArchitecture()
	Dim objWMIService, colItems, objItem
	Dim architecture

	Set objWMIService = GetWinMgmtsCIMV2()
	Set colItems = objWMIService.ExecQuery("SELECT AddressWidth FROM Win32_Processor")

	For Each objItem In colItems
		architecture = objItem.AddressWidth
	Next

	GetOSArchitecture = architecture
End Function

Function GetOSVersion()
	Dim objWMIService, colItems, objItem
	Dim osVersion, versionParts

	Set objWMIService = GetWinMgmtsCIMV2()
	Set colItems = objWMIService.ExecQuery("SELECT Version FROM Win32_OperatingSystem")

	For Each objItem In colItems
		osVersion = objItem.Version
	Next

	versionParts = Split(osVersion, ".")
	GetOSVersion = versionParts
End Function

Private Sub Imports(Files, FolderImports)
	If IsEmpty(FolderImports) Or FolderImports = "" Then FolderImports = "."
	
	Dim fso, file, filePath, fileObj
	Set fso = GetFileSystemObject()
	
	For Each file In Files
		If LCase(fso.GetExtensionName(file)) <> "vbs" Then file = file & ".vbs"
		filePath = fso.BuildPath(FolderImports, file)
		If fso.FileExists(filePath) Then
			Set fileObj = fso.OpenTextFile(filePath, 1)
			ExecuteGlobal fileObj.ReadAll
			fileObj.Close
		Else
			Log 1, "Arquivo " & file & " não encontrado na pasta " & FolderImports
		End If
	Next
End Sub

Function IsAdmin()
	Dim objShell, objExec, strOutput
	
	Set objShell = GetShell()
	Set objExec = objShell.Exec("Net Session")
	strOutput = objExec.StdOut.ReadAll
	
	IsAdmin = Len(strOutput) > 0
End Function

Function RunAs(ByVal Host)
	Dim oFSO, ScriptFolder, objShellApp

	Set objShellApp = GetShellApp()
	Set oFSO = GetFileSystemObject()

	ScriptFolder = oFSO.GetParentFolderName(WScript.ScriptFullName)
	objShellApp.ShellExecute Host, """" & WScript.ScriptFullName & """ " & GetArguments(), "", "RunAs", 1

End Function

Function SetRun(ByVal Host, ByVal Admin)
	If IsEmpty(Host) Or Host = "" Then Host = HOST_DEFAULT

	Dim needRunAs
	needRunAs = False

	If Admin And Not IsAdmin() Then needRunAs = True
	If Not InStr(LCase(WScript.FullName), LCase(Host)) > 0 Then needRunAs = True

	If needRunAs Then
		RunAs(Host)
		WScript.Quit
	End If

	SetRun = True
End Function

Function CreateGUID()
	Dim TypeLib
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	CreateGUID = Mid(TypeLib.Guid, 2, 36)
	Set TypeLib = Nothing
End Function

Function GetHeader(xmlhttp, headerName)
	On Error Resume Next
	Dim headerValue
	headerValue = xmlhttp.GetResponseHeader(headerName)
	If Err.Number <> 0 Then
		headerValue = ""
		Err.Clear
	End If
	On Error GoTo 0
	GetHeader = headerValue
End Function

Function DownloadFile(Url, SaveName, SaveFolder, LevelLog, ByRef SaveAs)
	Dim objHTTP, finalUrl, fileName, fileSize, fileDate, fso, fileExists, localFileSize, localFileDate, currentDir, contentType, fileExtension
	
	Set objHTTP = GetXMLHTTP()
	If objHTTP Is Nothing Then
		DownloadFile = "Erro ao criar objeto XMLHTTP."
		Exit Function
	End If
	
	Set fso = GetFileSystemObject()
	
	currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
	
	On Error GoTo 0
	objHTTP.Open "HEAD", Url, False
	objHTTP.Send
	
	If objHTTP.Status >= 300 And objHTTP.Status < 400 Then
		finalUrl = GetHeader(objHTTP, "Location")
		If finalUrl <> "" Then
			DownloadFile = DownloadFile(finalUrl, SaveName, SaveFolder, LevelLog, SaveAs)
			Exit Function
		End If
	End If
	
	If objHTTP.Status <> 200 Then
		DownloadFile = "Erro ao acessar URL: " & objHTTP.Status
		Exit Function
	End If
	
	contentType = GetHeader(objHTTP, "Content-Type")
	fileExtension = GetFileExtensionFromContentType(contentType)

	If IsEmpty(SaveName) Or SaveName = "" Then
		fileName = GetHeader(objHTTP, "Content-Disposition")
		If fileName = "" Then
			fileName = fso.BuildPath("", CreateGUID() & "." & IIf(fileExtension = "", "tmp", fileExtension))
		ElseIf InStr(fileName, "filename=") > 0 Then
			fileName = Mid(fileName, InStr(fileName, "filename=") + 9)
			fileName = Replace(fileName, """", "")
		End If
		SaveName = fileName
	Else
		If fileExtension <> "" Then
			'SaveName = fso.GetBaseName(SaveName) & "." & fileExtension
		End If
	End If
	
	fileSize = GetHeader(objHTTP, "Content-Length")
	If fileSize = "" Then fileSize = 0
	
	fileDate = GetHeader(objHTTP, "Last-Modified")
	If fileDate = "" Then fileDate = ""
	
	If IsEmpty(SaveFolder) Or SaveFolder = "" Then SaveFolder = currentDir
	
	SaveAs = fso.BuildPath(SaveFolder, SaveName)
	fileExists = fso.FileExists(SaveAs)
	
	If fileExists Then
		localFileSize = fso.GetFile(SaveAs).Size
		localFileDate = fso.GetFile(SaveAs).DateLastModified
	Else
		localFileSize = 0
		localFileDate = ""
	End If
	
	If (fileExists And localFileSize = fileSize) Or (fileExists And localFileDate = fileDate) Then
		DownloadFile = "Arquivo já baixado."
		Exit Function
	End If
	
	On Error GoTo 0
	objHTTP.Open "GET", Url, False
	objHTTP.Send
	
	If objHTTP.Status <> 200 Then
		DownloadFile = "Erro ao baixar o arquivo: " & objHTTP.Status
		Exit Function
	End If
	
	Dim stream
	Set stream = GetFileStream()
	
	stream.Type = 1
	stream.Open
	stream.Write objHTTP.responseBody
	stream.SaveToFile SaveAs, 2
	stream.Close
	
	DownloadFile = "Arquivo baixado com sucesso."
End Function

Sub Log(level, text)
	WScript.Echo "[" & Now & "]" & " [" & level & "] " & text
End Sub

Sub Main()
SetRun "", True

Dim apps(1)

apps(0) = Array("GLPI Agent", "GLPI-Agent.msi", 64, "/qn ADD_FIREWALL_EXCEPTION=1 ADDLOCAL=ALL AGENTMONITOR=1 RUNNOW=1 SYSTRAY=1 SERVER=https://gstic.patosdeminas.mg.gov.br/marketplace/glpiinventory /norestart FORCEUPDATE=1", "https://github.com/glpi-project/glpi-agent/releases/download/1.10/GLPI-Agent-1.10-x64.msi")
apps(1) = Array("RustDesk", "RustDesk.exe", 64, "--silent-install", "https://github.com/pmpatos/rustdesk/releases/download/nightly/rustdesk-1.3.0-x86_64.exe")

' Loop pelo array de aplicativos
Dim command, exitCode, objShell
Set objShell = GetShell()
	Dim app,description,param, url, SaveName, saveAs, savePath, osArchitecture, levelLog, fullPath

For Each app In apps
	description = app(0)
	SaveName = app(1)
	osArchitecture = app(2)
	param = app(3)
	url = app(4)
	exitCode=0
	savePath = ""
	levelLog = 1

	DownloadFile url, SaveName, savePath, levelLog, SaveAs

	If saveAs <> "" Then
		Log 2, "Instalando " & description & " "
	Else
		Log 1, "Falha no download: " & description
	End If

	command = SaveAs & " " & param
	exitCode = objShell.Run(command, 0, True)
	Log 1, Chr(9) & "Resultado da instalação: " & exitCode
Next
	
End Sub

Main
