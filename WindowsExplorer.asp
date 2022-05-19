<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Explorador de Archivos</TITLE>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			window.moveTo(100, 100);

			if (window.name) {
				if (window.name != 'WindowsExplorer') {
					window.location.replace('index.htm');
				}
			}
			else {
				window.location.replace('index.htm');
			}

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
				}
				return true;
			} // End of CheckLocalFile
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<FORM NAME="UploadLocalFileFrm" ID="UploadLocalFileFrm" ACTION="FileUploader.asp" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA" onSubmit="return CheckLocalFile(this)">
			<FONT FACE="Arial" SIZE="2"><BR />&nbsp;<B>Seleccione el archivo que contiene la información a subir:</B><BR /></FONT>
			&nbsp;<INPUT TYPE="FILE" NAME="LocalFile" ID="LocalFileFl" CLASS="TextFields" /><BR /><BR />
			<INPUT TYPE="HIDDEN" NAME="FolderName" ID="FolderNameHdn" VALUE="<%Response.Write oRequest("FolderName").Item%>" />
			<INPUT TYPE="HIDDEN" NAME="FileName" ID="FileNameHdn" VALUE="" />
			<INPUT TYPE="HIDDEN" NAME="Refresh" ID="RefreshHdn" VALUE="<%Response.Write oRequest("Refresh").Item%>" />
			&nbsp;<INPUT TYPE="SUBMIT" VALUE="Continuar" CLASS="Buttons" />
		</FORM>
	</BODY>
</HTML>
<%
Set oADODBConnection = Nothing
Call CleanFolderComponent(aFolderComponent)
%>