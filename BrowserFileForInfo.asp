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
<%
Dim sFolderName
Dim sFileName
Dim sRefresh
Dim sURL
Dim sNoOriginalFile
Dim sOriginalFileName

sOriginalFileName = CStr(oRequest("OriginalFile").Item)
sFolderName = UPLOADED_PHYSICAL_PATH

sFileName = oRequest("Action").Item & "_" & oRequest("UserID").Item & ".txt"

if oRequest("Action").Item = "ProcessForSar" then 
   sFileName = oRequest("Load").Item & "_" & oRequest("UserID").Item & ".txt"
end if

sRefresh = "1"
sURL = "BrowserFileForInfo.asp?FileReady=1"
sNoOriginalFile = ""
If Len(oRequest("NoOriginalFile").Item) > 0 Then sNoOriginalFile = "1"
If InStr(1, "DocsLibrary,Documents,Courses", oRequest("Action").Item, vbBinaryCompare) > 0 Then
	sFolderName = UPLOADED_PHYSICAL_PATH & oRequest("Action").Item & "\"
	sFileName = GetSerialNumberForDate("") & ".???"
	sURL = "BrowserFileForInfo.asp?Ready=1&Action=" & oRequest("Action").Item & "&NewFileName=<NEW_FILE_NAME />"
	sNoOriginalFile = "1"
End If
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE. Registro de información</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			<%If Len(oRequest("OriginalFile").Item) > 0 Then
				Response.Write "parent.window.location.href = parent.window.location.href + '&Step=2&OriginalFile=" & sOriginalFileName & "';"
			ElseIf Len(oRequest("FileReady").Item) > 0 Then
				Response.Write "parent.window.location.href = parent.window.location.href + '&Step=2';"
			End If%>
			function CheckLocalFile(oForm) {
				var iPosition = 0;
				var sFileName = '';

				if (oForm.LocalFile.value == '') {
					alert('Favor de seleccionar el archivo que desea subir.');
					oForm.LocalFile.focus();
					return false;
				}
				return true;
			} // End of CheckLocalFile
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If Len(oRequest("Ready").Item) = 0 Then%>
		<FORM NAME="UploadLocalFileFrm" ID="UploadLocalFileFrm" ACTION="FileUploader.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA" onSubmit="return CheckLocalFile(this)">
			<FONT FACE="Arial" SIZE="2"><B>Seleccione el archivo que contiene la información a subir:</B><BR /></FONT>
			<INPUT TYPE="FILE" NAME="LocalFile" ID="LocalFileFl" CLASS="TextFields" 
			/><INPUT TYPE="HIDDEN" NAME="FolderName" ID="FolderNameHdn" VALUE="<%Response.Write sFolderName%>"
			/><INPUT TYPE="HIDDEN" NAME="FileName" ID="FileNameHdn" VALUE="<%Response.Write sFileName%>"
			/><INPUT TYPE="HIDDEN" NAME="Refresh" ID="RefreshHdn" VALUE="<%Response.Write sRefresh%>"
			/><INPUT TYPE="HIDDEN" NAME="Action" ID="ActionHdn" VALUE="<%Response.Write oRequest("Action").Item%>"/>

            <% if oRequest("Action").Item = "ProcessForSar" then %>
            
            <INPUT TYPE="HIDDEN" NAME="Load" ID="LoadHdn" VALUE="<%Response.Write oRequest("Load").Item%>"
            /><INPUT TYPE="HIDDEN" NAME="UserId" ID="UserIdHdn" VALUE="<%Response.Write oRequest("UserId").Item%>"

            <% end if %>
			/><INPUT TYPE="HIDDEN" NAME="Step" ID="StepHdn" VALUE="2"
			/><INPUT TYPE="HIDDEN" NAME="URL" ID="URLHdn" VALUE="<%Response.Write sURL%>"
			/><INPUT TYPE="HIDDEN" NAME="NoOriginalFile" ID="NoOriginalFileHdn" VALUE="<%Response.Write sNoOriginalFile%>"
			/><IMG SRC="Images/Transparent.gif" WIDTH="30" HEIGHT="1" /><INPUT TYPE="SUBMIT" VALUE="Continuar" CLASS="Buttons"
		/></FORM>
		<%Else
			Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" HSPACE=""5"" /><FONT FACE=""Arial"" SIZE=""2""><B>El archivo está listo</B></FONT>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Select Case oRequest("Action").Item
					Case "Courses"
						Response.Write "if (parent.window.document.CourseFrm)" & vbNewLine
							Response.Write "parent.window.document.CourseFrm.CourseCertificate.value = '" & oRequest("NewFileName").Item & "';" & vbNewLine
					Case Else
						Response.Write "if (parent.window.document.CatalogFrm)" & vbNewLine
							Response.Write "parent.window.document.CatalogFrm.FilePath.value = '" & oRequest("NewFileName").Item & "';" & vbNewLine
				End Select
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If%>
	</BODY>
</HTML>