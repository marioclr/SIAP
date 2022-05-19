<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<%
Dim iIndex
Dim sTempFolderPath

If IsEmpty(oRequest("FolderPath").Item) Then
	aFolderComponent(S_PATH_FOLDER) = "."
	aFolderComponent(S_NAME_FOLDER) = "Uploaded Files"
End If

aFolderComponent(N_START_LEVEL_FOLDER) = 0
For iIndex = 1 To Len(SYSTEM_PHYSICAL_PATH)
	If StrComp(Mid(SYSTEM_PHYSICAL_PATH, iIndex, Len("\")), "\", vbBinaryCompare) = 0 Then
		aFolderComponent(N_START_LEVEL_FOLDER) = aFolderComponent(N_START_LEVEL_FOLDER) + 1
	End If
Next
aFolderComponent(N_DISPLAY_LEVEL_FOLDER) = aFolderComponent(N_START_LEVEL_FOLDER)
aFolderComponent(S_EXTRA_URL_FOLDER) = "&FormName=" & oRequest("FormName").Item & "&FieldName=" & oRequest("FieldName").Item

Call InitializeFolderComponent(oRequest, aFolderComponent)
sTempFolderPath = Replace(Replace((aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)), (SYSTEM_PHYSICAL_PATH & "Uploaded Files"), ""), "\", "\\")
If Len(sTempFolderPath) > 0 Then
	If StrComp(Left(sTempFolderPath, Len("\")), "\", vbBinaryCompare) = 0 Then
		sTempFolderPath = Right(sTempFolderPath, (Len(sTempFolderPath) - Len("\")))
	End If
End If
If Len(sTempFolderPath) > 0 Then
	If StrComp(Right(sTempFolderPath, Len("\")), "\", vbBinaryCompare) <> 0 Then
		sTempFolderPath = sTempFolderPath & "\\"
	End If
End If
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Explorador de Archivos</TITLE>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/HTMLLists.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/WindowsExplorer.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			if (window.name) {
				if (window.name != 'WindowsExplorer') {
					window.location.replace('index.htm');
				}
			}
			else {
				window.location.replace('index.htm');
			}

			var aFileNames = new Array(<%
				aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) = False
				aFolderComponent(B_DISPLAY_FILES_FOLDER) = True
				lErrorNumber = DisplayFolderContentsAsList(oRequest, aFolderComponent, sErrorDescription)
			%>'');

			function SendValueToOpener(oForm) {
				if (CheckFileSelection(oForm)) {
					window.opener.<%Response.Write oRequest("FormName").Item & "." & oRequest("FieldName").Item & ".value = '" & sTempFolderPath%>' + oForm.CourseFile.value;
					window.opener.<%Response.Write oRequest("FormName").Item & "." & oRequest("FieldName").Item & ".focus();"%>
					window.close();
				}
			} // End of SendValueToOpener
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<A NAME="TopAnchor"></A>
		<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
			<TR>
				<TD WIDTH="3" ROWSPAN="4"><IMG SRC="Images/Transparent.gif" WIDTH="3" HEIGHT="1" /></TD>
				<TD COLSPAN="4"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="3" /></TD>
				<TD WIDTH="3" ROWSPAN="4"><IMG SRC="Images/Transparent.gif" WIDTH="3" HEIGHT="1" /></TD>
			</TR>
			<TR><TD BGCOLOR="#000000" COLSPAN="4"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD></TR>
			<TR>
				<TD BGCOLOR="#000000" WIDTH="1"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
				<TD WIDTH="1"><NOBR>&nbsp;<IMG SRC="Images/IcnOpenFolder.gif" WIDTH="16" HEIGHT="16" />&nbsp;</NOBR></TD>
				<TD><FONT FACE="Verdana" SIZE="1">SIAP > <%lErrorNumber = DisplayFolderPath(oRequest, aFolderComponent, sErrorDescription)%></FONT></TD>
				<TD BGCOLOR="#000000" WIDTH="1"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
			</TR>
			<TR><TD BGCOLOR="#000000" COLSPAN="4"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD></TR>
		</TABLE>
		<FORM NAME="WindowsExplorerFrm" ID="WindowsExplorer" ACTION="FileExplorer.asp" METHOD="GET" onSubmit="return CheckFolderSelection(this)">
			<INPUT TYPE="HIDDEN" NAME="FolderPath" ID="FolderPathHdn" VALUE="<%Response.Write aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)%>" />
			<INPUT TYPE="HIDDEN" NAME="FolderStartLevel" ID="FolderStartLevelHdn" VALUE="<%Response.Write aFolderComponent(N_START_LEVEL_FOLDER)%>" />
			<INPUT TYPE="HIDDEN" NAME="FormName" ID="FormNameHdn" VALUE="<%Response.Write oRequest("FormName").Item%>" />
			<INPUT TYPE="HIDDEN" NAME="FieldName" ID="FieldNameHdn" VALUE="<%Response.Write oRequest("FieldName").Item%>" />
			<INPUT TYPE="HIDDEN" NAME="FilePath" ID="FilePathHdn" VALUE="<%Response.Write aFolderComponent(S_PATH_FOLDER)%>" />
			<INPUT TYPE="HIDDEN" NAME="FolderFilter" ID="FolderFilterHdn" VALUE="<%Response.Write aFolderComponent(S_FILTER_FOR_FILES_FOLDER)%>" />
			<TABLE WIDTH="98%" BORDER="0" CELLPADDING="0" CELLSPACING="0" ALIGN="CENTER">
				<TR>
					<TD VALIGN="TOP"><FONT FACE="Arial" SIZE="2">Directorios:</FONT><BR /><%
						aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) = True
						aFolderComponent(B_DISPLAY_FILES_FOLDER) = False
						Response.Write "<SELECT NAME=""FolderName"" ID=""FolderNameLst"" SIZE=""11"" CLASS=""Lists"" STYLE=""width: 200px"" onDblClick=""OpenFolder.click()"">" & vbNewLine
							lErrorNumber = DisplayFolderContentsInList(oRequest, aFolderComponent, sErrorDescription)
						Response.Write "</SELECT><FONT SIZE=""1""><BR /><BR /></FONT>" & vbNewLine
					%>
					</TD>
					<TD VALIGN="TOP"><FONT FACE="Arial" SIZE="2">Filtro:</FONT>
						<SELECT NAME="Filter" CLASS="Lists" onChange="ApplyFilter(aFileNames, this.value, document.WindowsExplorerFrm.CourseFile)">
							<OPTION VALUE="">Todos los archivos</OPTION>
							<%If Len(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) = 0 Then%>
								<OPTION VALUE=".doc">.doc</OPTION>
								<OPTION VALUE=".htm">.htm</OPTION>
								<OPTION VALUE=".asp">.pdf</OPTION>
								<OPTION VALUE=".ppt">.ppt</OPTION>
								<OPTION VALUE=".asp">.pps</OPTION>
								<OPTION VALUE=".txt">.txt</OPTION>
								<OPTION VALUE=".asp">.xls</OPTION>
							<%End If%>
						</SELECT><FONT SIZE="1"><BR /><BR /></FONT>
						<FONT FACE="Arial" SIZE="2">Archivos:</FONT><BR /><%
						aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) = False
						aFolderComponent(B_DISPLAY_FILES_FOLDER) = True
						Response.Write "<SELECT NAME=""CourseFile"" ID=""CourseFileLst"" SIZE=""9"" CLASS=""Lists"" STYLE=""width: 280px"" onDblClick=""SelectFile.click()"">" & vbNewLine
							lErrorNumber = DisplayFolderContentsInList(oRequest, aFolderComponent, sErrorDescription)
						Response.Write "</SELECT><FONT SIZE=""1""><BR /><BR /></FONT>" & vbNewLine
					%></TD>
				</TR>
				<TR>
					<TD ALIGN="LEFT">&nbsp;
						<INPUT TYPE="SUBMIT" NAME="OpenFolder" ID="OpenFolderBtn" VALUE="Abrir Directorio" CLASS="Buttons" />
					</TD>
					<TD ALIGN="RIGHT">
						<INPUT TYPE="BUTTON" NAME="SelectFile" ID="SelectFileBtn" VALUE="Seleccionar Archivo" CLASS="Buttons" onClick="SendValueToOpener(this.form)" />&nbsp;&nbsp;&nbsp;
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</BODY>
</HTML>
<%
Set oADODBConnection = Nothing
Call CleanFolderComponent(aFolderComponent)
%>