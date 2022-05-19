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
Dim sAction

sAction = oRequest("Action").Item
Select Case sAction
	Case "Rep_56"
		lErrorNumber = DeleteFile(Server.MapPath(REPORTS_PATH & "User_" & oRequest("UserID").Item & "\Rep_56_" & oRequest("UserID").Item & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sErrorDescription)
	Case "Reports"
		sErrorDescription = "No se pudo eliminar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Reports Where (UserID=" & oRequest("UserID").Item & ") And (ConstantID=" & oRequest("ConstantID").Item & ") And (ReportName='" & oRequest("ReportName").Item & "')", "Remove.asp", "_root", 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			lErrorNumber = DeleteFile(Server.MapPath(REPORTS_PATH & "User_" & oRequest("UserID").Item & "\Rep_" & oRequest("ConstantID").Item & "_" & oRequest("UserID").Item & "_" & oRequest("ReportName").Item & ".zip"), sErrorDescription)
		End If
End Select
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FONT FACE="Arial" SIZE="2">
			<%If lErrorNumber = 0 Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Select Case sAction
						Case "Rep_56"
							Response.Write "if (window.opener)" & vbNewLine
								Response.Write "window.opener.location.href='Budget.asp?Section=Money';" & vbNewLine
							Response.Write "window.close();" & vbNewLine
						Case Else
							Response.Write "if (window.opener)" & vbNewLine
								Response.Write "window.opener.location.reload();" & vbNewLine
							Response.Write "window.close();" & vbNewLine
					End Select
				Response.Write "//--></SCRIPT>" & vbNewLine
			Else
				Call DisplayErrorMessage("Error al eliminar el registro", sErrorDescription)
			End If%>
		</FONT>
	</BODY>
</HTML>