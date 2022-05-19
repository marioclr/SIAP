<%
Function DoLogout(oRequest, sAccessKey)
'************************************************************
'Purpose: To clean the connection cookies
'Inputs:  oRequest, sAccessKey
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoLogout"

	Response.Cookies("SIAP_AccessKey_" & sAccessKey) = ""
	Response.Cookies("SIAP_Password_" & sAccessKey) = ""
	Response.Cookies("SIAP_CurrentAccessKey") = ""
	Response.Cookies("SIAP_CurrentPassword") = ""
	Session.Contents("SIAP_CurrentAccessKey") = ""
	Session.Contents("SIAP_CurrentPassword") = ""

	DoLogout = Err.number
End function
%>