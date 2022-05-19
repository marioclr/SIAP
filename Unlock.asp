<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/FileLibrary.asp" -->
<!-- #include file="Libraries/EmailComponent.asp" -->
<%
Dim oRequest
Dim lErrorNumber
Dim sErrorDescription

Application.Contents("SIAP_Block") = ""
lErrorNumber = DeleteFile(Server.MapPath("Logs\Block.txt"), sErrorDescription)

aEmailComponent(S_TO_EMAIL) = "victor@jibda.com"
aEmailComponent(S_FROM_EMAIL) = "victor@jibda.com"
aEmailComponent(S_SUBJECT_EMAIL) = "SIAP ha sido desbloqueado"
aEmailComponent(S_BODY_EMAIL) = "<FONT FACE=""Arial"" SIZE=""2"">Este mensaje ha sido enviado por el Sistema de Administración del Personal (SIAP) ya que el sistema ha sido desbloqueado.<BR /></FONT>"
Call SendEmail(oRequest, aEmailComponent, sErrorDescription)

lErrorNumber = DeleteFile(Server.MapPath("Unlock.asp"), sErrorDescription)
%>