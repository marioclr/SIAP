<!-- #include file="Constants.asp" -->
<%
On Error Resume Next

Dim oRequest
Dim lErrorNumber
Dim sErrorDescription
Dim sAuxMessage
Dim oADODBConnection
Dim oADODBCommand
Dim param
Dim iConnectionType
Dim bIsNetscape
Dim bIsMac
Dim bWaitMessage
Dim iHelpSection
Dim iModuleID
Dim iGlobalSectionID
Dim lCurrentModule
Dim bTimeout

Dim iSADEConnectionType
Dim oSADEADODBConnection

If Request.Form.Count > 0 Then
	Set oRequest = Request.Form
Else
	Set oRequest = Request.QueryString
End If
lErrorNumber = 0
sErrorDescription = ""
%>
<!-- #include file="EmailComponent.asp" -->
<!-- #include file="GeneralLibrary.asp" -->
<%If CheckBlocked() Then Response.Redirect "Blocked.asp"%>
<!-- #include file="DatabaseLibrary.asp" -->
<!-- #include file="DaemonLibrary.asp" -->
<!-- #include file="DisplayLib.asp" -->
<!-- #include file="ErrorLogLib.asp" -->
<!-- #include file="FileLibrary.asp" -->
<!-- #include file="FolderComponent.asp" -->
<!-- #include file="QueriesLib.asp" -->
<!-- #include file="TablesLib.asp" -->
<%
iConnectionType = ORACLE
'iConnectionType = SQL_SERVER
iSADEConnectionType = SQL_SERVER
bIsNetscape = (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE", vbTextCompare) = 0)
bIsMac = (InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "Mac", vbTextCompare) > 0)
bWaitMessage = True
If Not IsEmpty(oRequest("HelpSection")) Then
	iHelpSection = CInt(oRequest("HelpSection").Item)
	iModuleID = CInt(oRequest("ModuleID").Item)
Else
	iHelpSection = 0
	iModuleID = -1
End If
If Not IsEmpty(oRequest("CurrentModuleID")) Then
	lCurrentModule = CInt(oRequest("CurrentModuleID").Item)
Else
	lCurrentModule = 1
End If
bTimeout = False

If B_SADE Then
	If InStr(1, ",ChangePassword.asp,Catalogs.asp,", ("," & GetASPFileName("") & ","), vbBinaryCompare) > 0 Then
		lErrorNumber = CreateADODBConnection(SADE_DATABASE_PATH, SADE_DATABASE_USERNAME, SADE_DATABASE_PASSWORD, iSADEConnectionType, oSADEADODBConnection, sErrorDescription)
		If (lErrorNumber <> 0) And (StrComp(GetASPFileName(""), "Default.asp", vbBinaryCompare) <> 0) Then
			Response.Redirect "Default.asp?SADE=1&ErrorID=" & L_ERR_NO_DB_CONNECTION
		End If
	End If
End If

lErrorNumber = CreateADODBConnection(SIAP_DATABASE_PATH, SIAP_DATABASE_USERNAME, SIAP_DATABASE_PASSWORD, iConnectionType, oADODBConnection, sErrorDescription)
If (lErrorNumber <> 0) And (StrComp(GetASPFileName(""), "Default.asp", vbBinaryCompare) <> 0) Then
	Response.Redirect "Default.asp?SIAP=1&ErrorID=" & L_ERR_NO_DB_CONNECTION
End If
%>
<!-- #include file="OptionsComponent.asp" -->
<!-- #include file="AlarmsLib.asp" -->
<!-- #include file="HeaderLib.asp" -->
<!-- #include file="MenuLib.asp" -->
<%
If lErrorNumber = 0 Then
	Call SendLogFileViaEmail()
	Call UpdateCurrenciesHistory(oRequest, oADODBConnection, "")
End If
iGlobalSectionID = CInt(Request.Cookies("SIAP_SectionID"))
%>