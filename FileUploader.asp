<%@LANGUAGE=VBSCRIPT%>
<!-- #include file="Libraries/Constants.asp" -->
<!-- #include file="Libraries/GeneralLibrary.asp" -->
<!-- #include file="Libraries/ErrorLogLib.asp" -->
<!-- #include file="Libraries/FileLibrary.asp" -->
<!-- #include file="Libraries/UploadLibrary.asp" -->
<!-- #include file="Libraries/PayrollResumeForSarComponent.asp" -->

<%
On Error Resume Next
Dim lErrorNumber
Dim sErrorDescription
Dim oFileUploader
Dim oFiles
Dim sFolderName
Dim sFileName
Dim sSpecificFolder

Dim sLoad
Dim sUserId
Dim sAction

Dim sTargetField
Dim sURL
Dim sRefresh
Dim bNoOriginalFile
Dim iTemp
Dim sOriginalFile
Dim alItems
Dim iIndex

lErrorNumber = 0
sErrorDescription = ""
Server.ScriptTimeOut = 72000
Set oFileUploader = New FileUploader
oFileUploader.Upload()

sTargetField = oFileUploader.GetFormValues("TargetField")
sURL = oFileUploader.GetFormValues("URL")
sRefresh = oFileUploader.GetFormValues("Refresh")
bNoOriginalFile = (Len(oFileUploader.GetFormValues("NoOriginalFile")) > 0)
sFolderName = oFileUploader.GetFormValues("FolderName")
sFileName = oFileUploader.GetFormValues("FileName")

sAction = oFileUploader.GetFormValues("Action")
sLoad   = oFileUploader.GetFormValues("Load")
sUserId = oFileUploader.GetFormValues("UserId")

sSpecificFolder = Left(sFolderName, Len(sFolderName) - (Len(sFolderName) - InStrRev(sFolderName, "\") + Len("\")))
sOriginalFile = oFileUploader.GetOriginalName()
alItems = Split(Replace(oFileUploader.GetFormValues("ItemIDs"), " ", ""), ",")
If InStr(1, ",QueryConsole.asp,QueriesConsole.asp,", "," & sURL & ",", vbBinaryCompare) = 0 Then sFolderName = SYSTEM_PHYSICAL_PATH & sFolderName
If StrComp(Right(sFolderName, Len("\")), "\", vbBinaryCompare) = 0 Then sFolderName = Left(sFolderName, (Len(sFolderName) - Len("\")))
Response.Flush
For Each oFiles In oFileUploader.oFiles.Items
	If Not FolderExists(sFolderName, sErrorDescription) Then
		lErrorNumber = CreateFolder(sFolderName, sErrorDescription)
		If lErrorNumber = 58 Then lErrorNumber = 0
	End If
	If lErrorNumber = 0 Then
		If Len(sFileName) > 0 Then
			If InStr(1, sFileName, ".???", vbBinaryCompare) > 0 Then sFileName = Replace(sFileName, ".???", Right(oFiles.sName, (Len(oFiles.sName) - InStrRev(oFiles.sName, ".") + Len("."))))
			oFiles.sName = sFileName
		End If
		Call oFiles.SaveFileAs(sFolderName, oFiles.sName, lErrorNumber, sErrorDescription)

		if sAction = "ProcessForSar" then FormatingTextTabColumns sLoad, sLoad & "_" & sUserId & ".txt"
		If InStr(1, sFolderName, "escaner") > 0 Then
            lErrorNumber = CopyFile(sFolderName & "\" & oFiles.sName, SYSTEM_PHYSICAL_PATH & sSpecificFolder & "\", sErrorDescription)
        End If

		sFileName = oFiles.sName
		If Err.number = 0 Then
			Call LogErrorInXMLFile(lErrorNumber, "Se agregó el archivo '" & sFolderName & "\" & oFiles.sName & "'.", 000, "FileUploader.asp", "_root", N_MESSAGE_LEVEL)
		End If
	End If
	If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
Next
If Not bNoOriginalFile Then
	If InStr(1, sURL, "?") > 0 Then
		sURL = sURL & "&OriginalFile=" & Server.URLEncode(sOriginalFile)
	Else
		sURL = sURL & "?OriginalFile=" & Server.URLEncode(sOriginalFile)
	End If
End If

Set oFileUploader = Nothing
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>File Uploader</TITLE>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="5" TOPMARGIN="5" MARGINWIDTH="5" MARGINHEIGHT="5" onLoad="window.focus()">
		<%If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al leer el archivo", sErrorDescription)
		ElseIf StrComp(sRefresh, "1", vbBinaryCompare) = 0 Then%>
			<SCRIPT LANGUAGE="JavaScript"><!--
				window.location.href = '<%Response.Write Replace(sURL, "<NEW_FILE_NAME />", sFileName)%>';
			//--></SCRIPT>
		<%ElseIf StrComp(sRefresh, "False", vbBinaryCompare) <> 0 Then%>
			<SCRIPT LANGUAGE="JavaScript"><!--
				<%If InStr(1, ",QueryConsole.asp,QueriesConsole.asp,", "," & GetASPFileName(sURL) & ",", vbBinaryCompare) = 0 Then%>
					if (window.opener) {
						<%If Len(Replace(RemoveParameterFromURLString(sURL, "OriginalFile"), "&", "")) > 0 Then%>
							window.opener.location.href = '<%Response.Write Replace(Replace(sURL, ";,;", "?"), ";;;", "&") & "&NewFile=" & oFiles.sName%>';
						<%Else%>
							window.opener.location.reload();
						<%End If%>
						window.opener.focus();
					}
					window.close();
				<%Else%>
					window.location.href = '<%Response.Write sURL%>';
				<%End If%>
			//--></SCRIPT>
		<%Else%>
			<SCRIPT LANGUAGE="JavaScript"><!--
				if (window.opener) {
					<%If Len(sTargetField) > 0 Then%>
						if (window.opener.window.document.<%Response.Write sTargetField%>)
							window.opener.window.document.<%Response.Write sTargetField%>.value = '<%Response.Write sFileName%>';
					<%End If%>
					window.opener.focus();
				}
				window.close();
			//--></SCRIPT>
		<%End If%>
	</BODY>
</HTML>