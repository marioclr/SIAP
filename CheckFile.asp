<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/DefaultLib.asp" -->
<%
Dim bFileReady
If InStr(1, oRequest("FileName").Item, ":", vbBinaryCompare) > 0 Then
	bFileReady = FileExists(oRequest("FileName").Item, sErrorDescription)
Else
	bFileReady = FileExists(Server.MapPath(oRequest("FileName").Item), sErrorDescription)
End If
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE. Reportes</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
		<%If (Not bFileReady) And (Len((oRequest("bNoReport").Item)) = 0) Then Response.Write "<META HTTP-EQUIV=""REFRESH"" CONTENT=""30; URL=CheckFile.asp?" & oRequest & """ />"%>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If Len(sErrorDescription) > 0 Then
			Call DisplayErrorMessage("Error al desplegar el reporte", sErrorDescription)
		ElseIf bFileReady Then
			sErrorDescription = "<IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""16"" HEIGHT=""16"" HSPACE=""3"" /><A HREF=""" & Replace(oRequest("FileName").Item, Server.MapPath(".") & "\", "") & """ TARGET=""Report"">Bajar un <B>archivo ZIP</B> con el <B>reporte generado</B></A>"
			Call DisplayInstructionsMessage("El reporte está listo", sErrorDescription)
		ElseIf Len((oRequest("bNoReport").Item)) > 0 Then
		Else%>
			<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR><TD ALIGN="CENTER">
				<IMG SRC="Images/AniWait.gif" WIDTH="100" HEIGHT="100" ALT="Cargando información..." /><BR /><BR />
				<FONT FACE="Arial" SIZE="2"><B>Cargando información...</B></FONT>
			</TD></TR></TABLE>
		<%End If%>
	</BODY>
</HTML>