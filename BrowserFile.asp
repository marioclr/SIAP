<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponentConstants.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/EmployeeSupportLib.asp" -->
<%
Dim sURL
Dim bLinks
Dim bIsClose
Dim sDocumentYear
Dim iYear

aOptionsComponent(L_ID_USER_OPTIONS) = aLoginComponent(N_USER_ID_LOGIN)
Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
If Len(oRequest("PaperworkID").Item) > 0 Then
	sDocumentYear = Left(oRequest("StartDate").Item, Len("0000"))
	sURL = "WindowsExplorer.asp?FolderName=" & Server.URLEncode(UPLOADED_PHYSICAL_PATH & "\escaner_" & sDocumentYear & "\\v" & CLng(oRequest("PaperworkNumber").Item))
	bIsClose = CBool(oRequest("isClose").Item)
Else
	sURL = "WindowsExplorer.asp?FolderName=" & Server.URLEncode(UPLOADED_PHYSICAL_PATH & "\e" & aEmployeeComponent(N_ID_EMPLOYEE))
End If
If Len(oRequest("RemoveFile").Item) > 0 Then
	If FileExists((oRequest("FilePath").Item & "\" & oRequest("FileName").Item), sErrorDescription) Then
		lErrorNumber = DeleteFile((oRequest("FilePath").Item & "\" & oRequest("FileName").Item), sErrorDescription)
		If lErrorNumber = 0 Then
			If Len(oRequest("PaperworkID").Item) > 0 Then
				lErrorNumber = GetPaperworksScannerYear(CStr(oRequest("FilePath").Item), iYear)
				If lErrorNumber = 0 Then
					If FileExists(SYSTEM_PHYSICAL_PATH & UPLOADED_PHYSICAL_PATH & "escaner_" & iYear & "\" & oRequest("FileName").Item, sErrorDescription) Then
						lErrorNumber = DeleteFile(SYSTEM_PHYSICAL_PATH & UPLOADED_PHYSICAL_PATH & "escaner_" & iYear & "\" & oRequest("FileName").Item, sErrorDescription)
					End If
				End If
				Response.Redirect "BrowserFile.asp?PaperworkID=" & CStr(oRequest("PaperworkID").Item) & "&PaperworkNumber=" & CStr(oRequest("PaperworkNumber").Item) & "&StartDate=" & CStr(oRequest("StartDate").Item) & "&isClose=" & CStr(oRequest("isClose").Item)
			Else
				Response.Redirect "BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN)
			End If
		End If
	End If
End If
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE. Archivos digitalizados</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FONT FACE="Arial" SIZE="2">
			<%If Len(oRequest("SecondFolder").Item) > 0 Then
				Response.Write "<B>Expediente adicional:</B><BR /><BR />"
			ElseIf Len(oRequest("Action").Item) = 0 And (Len(oRequest("PaperworkID").Item) > 0 And Not bIsClose) Then%>
				<B>Archivos digitalizados:</B><IMG SRC="Images/Transparent.gif" WIDTH="69" HEIGHT="1" />
				<IMG SRC="Images/IcnClip.gif" WIDTH="6" HEIGHT="10" />&nbsp;<A HREF="javascript: OpenNewWindow('<%Response.Write sURL%>', null, 'WindowsExplorer', 420, 80, 'yes', 'yes')" ID="AddFileLnk">Agregar archivo</A>&nbsp;
				<!-- #include file="Help_Files.asp" -->
            <%ElseIf Len(oRequest("Action").Item) = 0 Then%>
				<B>Archivos digitalizados:</B><IMG SRC="Images/Transparent.gif" WIDTH="69" HEIGHT="1" />
				<IMG SRC="Images/IcnClip.gif" WIDTH="6" HEIGHT="10" />&nbsp;<A HREF="javascript: OpenNewWindow('<%Response.Write sURL%>', null, 'WindowsExplorer', 420, 80, 'yes', 'yes')" ID="A1">Agregar archivo</A>&nbsp;
				<!-- #include file="Help_Files.asp" -->
			<%End If%>
			<DIV CLASS="SmallFolderContents"<%If Len(oRequest("Action").Item) > 0 Then Response.Write " STYLE=""height: 96px;"""%>><%
				aFolderComponent(S_PATH_FOLDER) = SYSTEM_PHYSICAL_PATH & UPLOADED_PHYSICAL_PATH
				bLinks = ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_FILES_PERMISSIONS) = N_DELETE_FILES_PERMISSIONS)
				Select Case oRequest("Action").Item
					Case "BanamexCensus", "SarCensus"
						aFolderComponent(S_NAME_FOLDER) = "\Census"
						aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = "Padron_"
						bLinks = False
						aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "window.parent.document.UploadFileInfoFrm.UploadFile.value = '<FILE_NAME>'; DoNothing();"	
					Case "ConsarFile"
						aFolderComponent(S_NAME_FOLDER) = "\Consar"
						aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = "Pago_"
						bLinks = False
						aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "window.parent.document.UploadFileInfoFrm.UploadFile.value = '<FILE_NAME>'; DoNothing();"
					Case "Filter"
						aFolderComponent(S_NAME_FOLDER) = "\Filters"
						aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = "Emp_"
						bLinks = False
						aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "window.parent.document.ReportFrm.EmployeeNumbers.value = '<FILE_NAME>'; DoNothing();"
					Case "Discos"
						aFolderComponent(S_NAME_FOLDER) = "\Discos"
						aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = "Dis_"
						bLinks = False
						aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "window.parent.document.UploadFileInfoFrm.UploadFile.value = '<FILE_NAME>'; DoNothing();"
					Case "Prestaciones"
						aFolderComponent(S_NAME_FOLDER) = "\Prestaciones"
						aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = "Pre_"
						bLinks = False
						aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "window.parent.document.UploadFileInfoFrm.UploadFile.value = '<FILE_NAME>'; DoNothing();"
					Case Else
						If Len(oRequest("SecondFolder").Item) > 0 Then
							sURL = Replace(Replace(SECOND_PATH & "<FOLDER_NAME>/<FILE_NAME>", "\", "\\"), "/", "\/")
							aFolderComponent(S_PATH_FOLDER) = SECOND_PHYSICAL_PATH
							aFolderComponent(S_NAME_FOLDER) = "\e" & aEmployeeComponent(N_ID_EMPLOYEE)
							aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "OpenFileInNewWindow('" & sURL & "', null, 'WindowsExplorer');"
							bLinks = False
						ElseIf Len(oRequest("PaperworkID").Item) > 0 Then
							lErrorNumber = CopyPaperworkFiles(oRequest, CLng(Left(oRequest("StartDate").Item, Len("0000"))), CLng(oRequest("PaperworkNumber").Item), sErrorDescription)
							If bLinks Then
								bLinks = Not bIsClose
							End If
							If Not bLinks Then
								bLinks = InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0
							End If
							sURL = Replace(Replace(UPLOADED_PHYSICAL_PATH & "<FOLDER_NAME>/<FILE_NAME>", "\", "\\"), "/", "\/")
							aFolderComponent(S_NAME_FOLDER) = "\escaner_" & sDocumentYear & "\v" & CLng(oRequest("PaperworkNumber").Item)
							aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "OpenFileInNewWindow('" & sURL & "', null, 'WindowsExplorer');"
						Else
							sURL = Replace(Replace(UPLOADED_PHYSICAL_PATH & "<FOLDER_NAME>/<FILE_NAME>", "\", "\\"), "/", "\/")
							aFolderComponent(S_NAME_FOLDER) = "\e" & aEmployeeComponent(N_ID_EMPLOYEE)
							aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = "OpenFileInNewWindow('" & sURL & "', null, 'WindowsExplorer');"
						End If
				End Select
				If FolderExists((aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)), sErrorDescription) Then
					lErrorNumber = DisplayFolderContents(oRequest, bLinks, aFolderComponent, sErrorDescription)
					If Len(oRequest("Action").Item) = 0 Then
						If aFolderComponent(B_IS_EMPTY_FOLDER) Then
							lErrorNumber = DeleteFolder(Server.MapPath(UPLOADED_PHYSICAL_PATH & "e" & aEmployeeComponent(N_ID_EMPLOYEE)), sErrorDescription)
						End If
					End If
				Else
					Response.Write "<BR />&nbsp;No se han agregado archivos digitalizados."
				End If
			%></DIV>
		</FONT>
	</BODY>
</HTML>