<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 3000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/LoginComponentConstants.asp" -->
<%
Dim sFileContents
Dim sQuery
Dim asQuery
Dim asResults
Dim iIndex
Dim oRecordset
Dim bFile

'If Not aLoginComponent(B_ADMINISTRATOR_LOGIN) And Not aLoginComponent(B_TUTOR_LOGIN) Then
'	Response.Redirect REDIRECT_INVALID_USER_PAGE & "?Permission=0"
'End If

If Len(oRequest("FolderName").Item) = 0 Then
	aFolderComponent(S_PATH_FOLDER) = SYSTEM_PHYSICAL_PATH
	aFolderComponent(S_NAME_FOLDER) = "."
End If
Call InitializeFolderComponent(oRequest, aFolderComponent)
If Len(oRequest("RemoveFile").Item) > 0 Then
	If FileExists(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER) & "\" & oRequest("FileName").Item, sErrorDescription) Then
		Call DeleteFile(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER) & "\" & oRequest("FileName").Item, sErrorDescription)
	End If
End If

If (Len(oRequest("ApplicationContents").Item) > 0) And (Len(oRequest("a4").Item) > 0) Then
	'If CInt(oRequest("a4").Item) = Minute(Time()) Then
		Application.Contents(oRequest("VariableName").Item) = oRequest("VariableValue").Item
	'End If
End If
If (Len(oRequest("CopyLog").Item) > 0) And (Len(oRequest("a0").Item) > 0) Then
	If CInt(oRequest("a0").Item) = Minute(Time()) Then
		If Len(oRequest("UndoCopy").Item) = 0 Then
			lErrorNumber = CopyFile(Server.MapPath("Logs\Log" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), Server.MapPath("Logs\Log" & Left(GetSerialNumberForDate(""), Len("00000000")) & "_.txt"), sErrorDescription)
		Else
			lErrorNumber = CopyFile(Server.MapPath("Logs\Log" & Left(GetSerialNumberForDate(""), Len("00000000")) & "_.txt"), Server.MapPath("Logs\Log" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sErrorDescription)
			If lErrorNumber = 0 Then lErrorNumber = DeleteFile(Server.MapPath("Logs\Log" & Left(GetSerialNumberForDate(""), Len("00000000")) & "_.txt"), sErrorDescription)
		End If
	End If
End If
If (Len(oRequest("SendEmail").Item) > 0) And (Len(oRequest("a1").Item) > 0) Then
	If CInt(oRequest("a1").Item) = Minute(Time()) Then
		Response.Write "Ok<BR />"
		aEmailComponent(S_FROM_EMAIL) = S_ADMIN_EMAIL_ACCOUNT
		aEmailComponent(S_TO_EMAIL) = "victor@jibda.com"
		aEmailComponent(S_SUBJECT_EMAIL) = "Prueba"
		aEmailComponent(S_BODY_EMAIL) = "<FONT FACE=""Arial"" SIZE=""2"">Este mensaje ha sido enviado desde QueryConsole.asp como una prueba</FONT>"
		lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
		Response.Write lErrorNumber
	End If
End If
If (Len(oRequest("SendFile").Item) > 0) And (Len(oRequest("a2").Item) > 0) Then
	If CInt(oRequest("a2").Item) = Minute(Time()) Then
		Response.Write "Ok<BR />"
		If Len(oRequest("SaveFile").Item) > 0 Then
			lErrorNumber = SaveTextToFile(oRequest("FilePath").Item, oRequest("FileContents").Item, sErrorDescription)
		Else
			sFileContents = GetFileContents(oRequest("FilePath").Item, sErrorDescription)
		End If
	End If
End If
If (Len(oRequest("CreateFolder").Item) > 0) And (Len(oRequest("a3").Item) > 0) Then
	If CInt(oRequest("a3").Item) = Minute(Time()) Then
		Response.Write "Ok<BR />"
		lErrorNumber = CreateFolder(oRequest("FolderPath").Item, sErrorDescription)
	End If
End If
aHeaderComponent(S_TITLE_NAME_HEADER) = "Consola de Queries"
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE><%Response.Write aHeaderComponent(S_WINDOW_TITLE_HEADER)%></TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/RollOver.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			function CheckLocalFile(oForm) {
				var iPosition = 0;
				var sFileName = '';

				if (oForm.LocalFile.value == '') {
					alert('Favor de seleccionar el archivo que desea subir.');
					oForm.LocalFile.focus();
					return false;
				}
				else {
					sFileName = oForm.LocalFile.value;
					do {
						iPosition = sFileName.search(/\\/gi);
						if (iPosition > -1)
							iPosition++;
						sFileName = sFileName.substr(iPosition);
					}
					while (iPosition > -1);
					oForm.FileName.value = sFileName;
					<%If Len(oRequest("TargetField").Item) > 0 Then%>
						if (opener.window.document.<%Response.Write oRequest("TargetField").Item%>)
							opener.window.document.<%Response.Write oRequest("TargetField").Item%>.value = sFileName;
					<%End If%>
				}
				return true;
			} // End of CheckLocalFile
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- BEGIN: CONTENTS -->
		<FONT FACE="Arial" SIZE="2"><BR /><BR />Usted se encuentra aquí:&nbsp;<A HREF="Main.asp">Inicio</A> > <A HREF="Tools.asp">Herramientas</A> > <B>Consola de queries</B><BR /><BR /></FONT>
		<%If lErrorNumber <> 0 Then
			If Len(oRequest("SendEmail").Item) > 0 Then
				sErrorDescription = sErrorDescription & "<BR />Server: " & aEmailComponent(S_SERVER_NAME_EMAIL) & "<BR />To: " & aEmailComponent(S_TO_EMAIL) & "<BR />Cc: " & aEmailComponent(S_CC_EMAIL) & "<BR />Bcc: " & aEmailComponent(S_BCC_EMAIL) & "<BR />From: " & aEmailComponent(S_FROM_EMAIL) & "<BR />Subject: " & aEmailComponent(S_SUBJECT_EMAIL) & "<BR />Body: " & aEmailComponent(S_BODY_EMAIL)
				Call DisplayErrorMessage("Ocurrió un error al enviar el mensaje", sErrorDescription)
			ElseIf Len(oRequest("SendFile").Item) > 0 Then
				sErrorDescription = sErrorDescription & "<BR />File: " & oRequest("FilePath").Item
				Call DisplayErrorMessage("Ocurrió un error con el archivo", sErrorDescription)
			Else
				Call DisplayErrorMessage("Ocurrió un error al ejecutar el query sobre la base de datos", sErrorDescription)
			End If
			lErrorNumber = 0
			sErrorDescription = ""
		End If%>
		<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR>
			<TD WIDTH="50%" VALIGN="TOP">
				<FORM NAME="QueryConsoleFrm" ID="QueryConsoleFrm" ACTION="QueriesConsole.asp" METHOD="POST">
					&nbsp;<FONT FACE="Arial" SIZE="2">
						Introducir el query a ejecutar:
						<IMG SRC="Images/Transparent.gif" WIDTH="80" HEIGHT="1" />
						<INPUT TYPE="CHECKBOX" NAME="AsCount" ID="AsCountChk" VALUE="1" />Count&nbsp;&nbsp;&nbsp;
						<INPUT TYPE="CHECKBOX" NAME="AsText" ID="AsTextChk" VALUE="1" />Como texto
					</FONT><BR />
					&nbsp;<TEXTAREA NAME="Query" ID="QueryTxtArea" ROWS="5" COLS="50" CLASS="TextFields"><%Response.Write sQuery%></TEXTAREA>
					<BR /><BR />
					&nbsp;<INPUT TYPE="SUBMIT" VALUE="Ejecutar Query" CLASS="Buttons" />
					<IMG SRC="Images/Transparent.gif" WIDTH="190" HEIGHT="1" />
					<A HREF="javascript: ToggleDisplay(document.all['ApplicationContentsDiv'])"><IMG SRC="Images/Transparent.gif" WIDTH="10" HEIGHT="10" BORDER="0" ALT="Application.Contents" /></A>
					<A HREF="javascript: ToggleDisplay(document.all['CopyLogDiv'])"><IMG SRC="Images/Transparent.gif" WIDTH="10" HEIGHT="10" BORDER="0" ALT="Log" /></A>
					<A HREF="javascript: ToggleDisplay(document.all['SendEmailDiv'])"><IMG SRC="Images/Transparent.gif" WIDTH="10" HEIGHT="10" BORDER="0" ALT="Test e-mail" /></A>
					<A HREF="javascript: ToggleDisplay(document.all['FileDiv'])"><IMG SRC="Images/Transparent.gif" WIDTH="10" HEIGHT="10" BORDER="0" ALT="Files" /></A>
				</FORM>
				<DIV NAME="ApplicationContentsDiv" ID="ApplicationContentsDiv" STYLE="display: none">
					<FORM NAME="ApplicationContentsFrm" ID="ApplicationContentsFrm" ACTION="QueriesConsole.asp" METHOD="GET">
						<FONT FACE="Arial" SIZE="2"><%Response.Write Time()%></FONT>
						<INPUT TYPE="TEXT" NAME="a4" ID="a4Txt" SIZE="2" MAXLENGTH="2" VALUE="" CLASS="TextFields" /><BR />
						Nombre: <INPUT TYPE="TEXT" NAME="VariableName" ID="VariableNameTxt" SIZE="20" MAXLENGTH="100" VALUE="" CLASS="TextFields" /><BR />
						Valor: <INPUT TYPE="TEXT" NAME="VariableValue" ID="VariableValueTxt" SIZE="100" VALUE="" CLASS="TextFields" /><BR />
						<INPUT TYPE="SUBMIT" NAME="ApplicationContents" ID="ApplicationContentsBtn" VALUE="Cambiar Valor" CLASS="Buttons" />
					</FORM>
				</DIV>
				<DIV NAME="CopyLogDiv" ID="CopyLogDiv" STYLE="display: none">
					<FORM NAME="CopyLogFrm" ID="CopyLogFrm" ACTION="QueriesConsole.asp" METHOD="GET">
						<FONT FACE="Arial" SIZE="2"><%Response.Write Time()%></FONT>
						<INPUT TYPE="TEXT" NAME="a0" ID="a0Txt" SIZE="2" MAXLENGTH="2" VALUE="" CLASS="TextFields" />
						<INPUT TYPE="CHECKBOX" NAME="UndoCopy" ID="UndoCopyChk" VALUE="1" />
						<INPUT TYPE="SUBMIT" NAME="CopyLog" ID="CopyLogBtn" VALUE="Duplicar Bitácora" CLASS="Buttons" />
					</FORM>
				</DIV>
				<DIV NAME="SendEmailDiv" ID="SendEmailDiv" STYLE="display: none">
					<FORM NAME="SendEmailFrm" ID="SendEmailFrm" ACTION="QueriesConsole.asp" METHOD="GET">
						<FONT FACE="Arial" SIZE="2"><%Response.Write Time()%></FONT>
						<INPUT TYPE="TEXT" NAME="a1" ID="a1Txt" SIZE="2" MAXLENGTH="2" VALUE="" CLASS="TextFields" />
						<INPUT TYPE="SUBMIT" NAME="SendEmail" ID="SendEmailBtn" VALUE="Enviar" CLASS="Buttons" />
					</FORM>
				</DIV>
				<DIV NAME="FileDiv" ID="FileDiv" STYLE="<%If Not bIsMac Then Response.Write "display: none"%>">
					<FORM NAME="FileFrm" ID="FileFrm" ACTION="QueriesConsole.asp" METHOD="POST">
						<TEXTAREA NAME="FileContents" ID="FileContentsTxtArea" ROWS="5" COLS="40" CLASS="TextFields"><%Response.Write sFileContents%></TEXTAREA><BR />
						<INPUT TYPE="TEXT" NAME="FilePath" ID="FilePathTxt" SIZE="40" VALUE="<%Response.Write oRequest("FilePath").Item%>" CLASS="TextFields" /><BR />
						<FONT FACE="Arial" SIZE="2">
							<INPUT TYPE="CHECKBOX" NAME="SaveFile" ID="SaveFileChk" VALUE="1" /> Save<BR />
							<%Response.Write Time()%>
							<INPUT TYPE="TEXT" NAME="a2" ID="a2Txt" SIZE="2" MAXLENGTH="2" VALUE="" CLASS="TextFields" />
							<INPUT TYPE="SUBMIT" NAME="SendFile" ID="SendFileBtn" VALUE="Enviar" CLASS="Buttons" /><BR />
							<%Response.Write aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)%>
						</FONT>
						<DIV CLASS="FolderContents" STYLE="width: 100%">
							<%lErrorNumber = DisplayFolderContents(oRequest, True, aFolderComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								Call DisplayErrorMessage("Error en el directorio", sErrorDescription)
							End If%>
						</DIV>
					</FORM>
					<FORM NAME="UploadLocalFileFrm" ID="UploadLocalFileFrm" ACTION="FileUploader.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA" onSubmit="return CheckLocalFile(this)"><FONT FACE="Arial" SIZE="2">
						&nbsp;<B>Seleccione el archivo que contiene la información a subir:</B><BR />
						&nbsp;<INPUT TYPE="FILE" NAME="LocalFile" ID="LocalFileFl" CLASS="TextFields" /><BR />
						&nbsp;Ruta donde se pondrá el archivo: <INPUT TYPE="TEXT" NAME="FolderName" ID="FolderNameHdn" VALUE="<%Response.Write oRequest("FolderName").Item%>" SIZE="30" CLASS="TextFields" /><BR />
						<INPUT TYPE="HIDDEN" NAME="FileName" ID="FileNameHdn" VALUE="" />
						<INPUT TYPE="HIDDEN" NAME="URL" ID="URLHdn" VALUE="QueriesConsole.asp" />
						&nbsp;<INPUT TYPE="SUBMIT" VALUE="Continuar" CLASS="Buttons" / id=SUBMIT1 name=SUBMIT1>
					</FONT></FORM>
				</DIV>
				<DIV NAME="FolderDiv" ID="FolderDiv" STYLE="<%If Not bIsMac Then Response.Write "display: none"%>">
					<FORM NAME="FolderFrm" ID="FolderFrm" ACTION="QueriesConsole.asp" METHOD="POST">
						<INPUT TYPE="TEXT" NAME="FolderPath" ID="FolderPathTxt" SIZE="40" VALUE="<%Response.Write oRequest("FolderPath").Item%>" CLASS="TextFields" /><BR />
						<FONT FACE="Arial" SIZE="2">
							<%Response.Write Time()%>
							<INPUT TYPE="TEXT" NAME="a3" ID="a3Txt" SIZE="2" MAXLENGTH="2" VALUE="" CLASS="TextFields" />
							<INPUT TYPE="SUBMIT" NAME="CreateFolder" ID="CreateFolderBtn" VALUE="Crear" CLASS="Buttons" /><BR />
							<%Response.Write aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)%>
						</FONT>
					</FORM>
				</DIV>
			</TD>
			<TD WIDTH="50%" VALIGN="TOP"><FONT FACE="Arial" SIZE="2"><DIV STYLE="width: 490px; height: 200px; overflow: auto; border: 1pt solid #000000;">
				<%lErrorNumber = ShowADODBConnectionProperties(oADODBConnection, sErrorDescription)
				If lErrorNumber <> 0 Then
					Call DisplayErrorMessage("Ocurrió un error al mostrar las propiedades de la conexión", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
				Response.Write "<BR />"
				Call TestFileSystemObject(Server.MapPath("Styles"), sErrorDescription)%>
			</DIV></FONT></TD>
		</TR></TABLE>
		<BR />
		<FONT FACE="Arial" SIZE="2">
			<%Response.Flush()
			sQuery = oRequest("Query").Item
			If StrComp(Left(sQuery, Len("*")), "*", vbBinaryCompare) = 0 Then
				sQuery = Replace(sQuery, "*", "", 1, 1, vbBinaryCompare)
				bFile = False
				If InStr(1, sQuery, ":", vbBinaryCompare) = 2 Then
					If FileExists(sQuery, sErrorDescription) Then
						sQuery = GetFileContents(sQuery, sErrorDescription)
						bFile = True
					Else
						sQuery = ""
					End If
				End If
				asQuery = Split(sQuery, (";" & vbNewLine), -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asQuery)
					asQuery(iIndex) = Trim(Replace(asQuery(iIndex), vbNewLine, " ", 1, -1, vbBinaryCompare))
					If Len(asQuery(iIndex)) > 0 Then
						sErrorDescription = "Ocurrió un error al ejecutar el query sobre la base de datos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, asQuery(iIndex), "QueryConsole.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Ocurrió un error al ejecutar el query sobre la base de datos", sErrorDescription)
							asQuery(iIndex) = asQuery(iIndex) & " (error)"
							lErrorNumber = 0
							sErrorDescription = ""
						Else
							If (Len(oRequest("AsCount").Item) > 0) And (InStr(1, asQuery(iIndex), "Select Count(*) From ", vbBinaryCompare) > 0) Then
								asResults = asResults & asQuery(iIndex) & SECOND_LIST_SEPARATOR & oRecordset.Fields(0).Value & LIST_SEPARATOR
							ElseIf Not bFile Then
								Response.Write "<B>" & CleanStringForHTML(asQuery(iIndex)) & "</B><BR /><BR />"
								If (InStr(1, Trim(asQuery(iIndex)), "Select ", vbTextCompare) = 1) Or (InStr(1, sQuery, " ", vbBinaryCompare) = 0) Then Call DisplayRecordsetAsTable(oRecordset, "", "", "", "", (Len(oRequest("AsText").Item) > 0), sErrorDescription)
								Response.Write "<BR /><HR WIDTH=""98%"" /><BR />"
							End If
						End If
					End If
				Next
				If (Len(oRequest("AsCount").Item) > 0) And (Len(asResults) > 0) Then
					asResults = Split(asResults, LIST_SEPARATOR)
					Response.Write vbNewLine & vbNewLine
					For iIndex = 0 To UBound(asResults) - 1
						asResults(iIndex) = Split(asResults(iIndex), SECOND_LIST_SEPARATOR)
						Response.Write Replace(Replace(asResults(iIndex)(0), "Select Count(*) From ", ""), ";", "") & vbTab
						Response.Write asResults(iIndex)(1) & vbNewLine
					Next
					Response.Write vbNewLine & vbNewLine
				End If
				If (Len(sQuery) > 0) And (Not bFile) Then
					Response.Write "&nbsp;<B>Queries ejecutados:</B><BR />"
					For iIndex = 0 To UBound(asQuery)
						If Len(asQuery(iIndex)) > 0 Then
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(asQuery(iIndex)) & "<BR />"
						End If
					Next
				End If
			End If
			If Len(sQuery) = 0 Then
				Call LogErrorInXMLFile(0, "El usuario entro a la consola de queries", 000, "_root", "_root", N_MESSAGE_LEVEL)
			End If%>
		</FONT><BR /><BR /><BR />
		<!-- END: CONTENTS -->
	</BODY>
</HTML>
<%
oRecordset.Close
Set oRecordset = Nothing
Set oADODBConnection = Nothing
%>