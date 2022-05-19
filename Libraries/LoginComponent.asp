<!-- #include file="LogoutLib.asp" -->
<!-- #include file="LoginComponentConstants.asp" -->
<%
Dim bCheckUserLoginOnly
bCheckUserLoginOnly = True
If (Len(oRequest("DisplayError").Item) = 0) Or IsEmpty(oRequest("DisplayError").Item) Then
	If lErrorNumber = 0 Then
		lErrorNumber = GetLoginCredencials(oRequest, oADODBConnection, aLoginComponent, sErrorDescription)
		If lErrorNumber <> 0 Then
			Call DoLogout(oRequest, aLoginComponent(S_ACCESS_KEY_LOGIN))
		Else
			If aLoginComponent(N_USER_ID_LOGIN) <> -1 Then
				aOptionsComponent(L_ID_USER_OPTIONS) = aLoginComponent(N_USER_ID_LOGIN)
				Call GetOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
			End If
			lErrorNumber = RedirectUser(oRequest, oADODBConnection, aLoginComponent, aOptionsComponent, lErrorNumber, sErrorDescription)
		End If
	End If
End If

Function InitializeLoginComponent(oRequest, aLoginComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Login Component
'         using Cookies, the URL parameters or default values
'Inputs:  oRequest
'Outputs: aLoginComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeLoginComponent"
	Dim aTemp
	Dim iIndex
	Redim Preserve aLoginComponent(N_LOGIN_COMPONENT_SIZE)

	If IsEmpty(aLoginComponent(S_ACCESS_KEY_LOGIN)) Then
		If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then
			aLoginComponent(S_ACCESS_KEY_LOGIN) = Request.Cookies("SIAP_CurrentAccessKey")
			If Len(aLoginComponent(S_ACCESS_KEY_LOGIN)) = 0 Then
				aLoginComponent(S_ACCESS_KEY_LOGIN) = Session.Contents("SIAP_CurrentAccessKey")
			End If
		End If
		If (Len(aLoginComponent(S_ACCESS_KEY_LOGIN)) = 0) Or (Len(oRequest("Login").Item) > 0) Then
			aLoginComponent(S_ACCESS_KEY_LOGIN) = oRequest("AccessKey").Item
		End If
	End If
	If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then
		Response.Cookies("SIAP_CurrentAccessKey") = aLoginComponent(S_ACCESS_KEY_LOGIN)
		Response.Cookies("SIAP_AccessKey_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = aLoginComponent(S_ACCESS_KEY_LOGIN)
	End If

	If IsEmpty(aLoginComponent(S_PASSWORD_LOGIN)) Then
		If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then
			aLoginComponent(S_PASSWORD_LOGIN) = Request.Cookies("SIAP_Password_" & aLoginComponent(S_ACCESS_KEY_LOGIN)).Item
			If Len(aLoginComponent(S_PASSWORD_LOGIN)) = 0 Then
				aLoginComponent(S_PASSWORD_LOGIN) = Request.Cookies("SIAP_CurrentPassword")
				If Len(aLoginComponent(S_PASSWORD_LOGIN)) = 0 Then
					aLoginComponent(S_PASSWORD_LOGIN) = Session.Contents("SIAP_CurrentPassword")
				End If
			End If
		End If
		If (Len(aLoginComponent(S_PASSWORD_LOGIN)) = 0) Or (Len(oRequest("Login").Item) > 0) Then
			aLoginComponent(S_PASSWORD_LOGIN) = oRequest("Password").Item
		End If
	End If
	If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then
		Response.Cookies("SIAP_CurrentPassword") = aLoginComponent(S_PASSWORD_LOGIN)
		Response.Cookies("SIAP_Password_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = aLoginComponent(S_PASSWORD_LOGIN)
	End If

	aLoginComponent(B_VALID_USER_LOGIN) = False
	aLoginComponent(B_EXPIRED_LOGIN) = False
	aLoginComponent(B_BLOCKED_LOGIN) = False
	aLoginComponent(N_USER_ID_LOGIN) = -1
	aLoginComponent(S_USER_NAME_LOGIN) = ""
	aLoginComponent(S_USER_LAST_NAME_LOGIN) = ""
	aLoginComponent(N_PROFILE_ID_LOGIN) = -1
	aLoginComponent(N_USER_PERMISSIONS_LOGIN) = 0
	aLoginComponent(N_USER_PERMISSIONS2_LOGIN) = 0
	aLoginComponent(N_USER_PERMISSIONS3_LOGIN) = 0
	aLoginComponent(N_USER_PERMISSIONS4_LOGIN) = 0
	aLoginComponent(L_PERMISSION_REPORTS_LOGIN) = 0
	aLoginComponent(L_PERMISSION_REPORTS2_LOGIN) = 0
	aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) = "-1"
	aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) = -1
	aLoginComponent(S_USER_E_MAIL_LOGIN) = ""
	aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN) = ""
	aLoginComponent(S_EMPLOYEE_NUMBER_LOGIN) = ""
	aLoginComponent(B_ACTIVE_LOGIN) = True
	aLoginComponent(B_TECH_SUPPORT_LOGIN) = False
	aLoginComponent(B_CUSTOMER_LOGIN) = True

	aLoginComponent(S_PATH_LOGIN) = ""
	aTemp = Split(Right(Request.ServerVariables("PATH_INFO"), Len(Request.ServerVariables("PATH_INFO")) - Len(SYSTEM_PATH)), "/", -1, vbBinaryCompare)
	For iIndex=0 To (UBound(aTemp) - 1)
		aLoginComponent(S_PATH_LOGIN) = aLoginComponent(S_PATH_LOGIN) & "../"
	Next

	aLoginComponent(B_COMPONENT_INITIALIZED_LOGIN) = True
	InitializeLoginComponent = Err.number
	Err.Clear
End Function

Function GetLoginCredencials(oRequest, oADODBConnection, aLoginComponent, sErrorDescription)
'************************************************************
'Purpose: To get the credentials from the database using the
'         access key and password
'Inputs:  oRequest, oADODBConnection, aLoginComponent
'Outputs: aLoginComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLoginCredencials"
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aLoginComponent(B_COMPONENT_INITIALIZED_LOGIN)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeLoginComponent(oRequest, aLoginComponent)
	End If
	If lErrorNumber = 0 Then
		If Len(aLoginComponent(S_ACCESS_KEY_LOGIN)) > 0 Then
			If Not B_PORTAL Then sCondition = " And (UserPassword='" & Replace(aLoginComponent(S_PASSWORD_LOGIN), "'", "") & "')"
			sErrorDescription = "No se pudieron obtener las credenciales del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Users.*, ChangeDate, ZonePath From Users, UsersPWD, Zones Where (Users.UserID=UsersPWD.UserID) And (Users.PermissionZoneID=Zones.ZoneID) And (UserAccessKey='" & Replace(aLoginComponent(S_ACCESS_KEY_LOGIN), "'", "") & "')" & sCondition, "LoginComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					Call CheckLoginFailures(oADODBConnection, aLoginComponent, True)
					lErrorNumber = L_ERR_INCORRECT_PASSWORD
					sErrorDescription = "La clave de acceso y/o la contraseña <FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>están incorrectos</B></FONT>. Favor de introducirlos de nuevo. Recuerde que esta información <FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>es sensible a mayúsculas y minúsculas</B></FONT>."
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "LoginComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
					If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then Response.Cookies("SIAP_AccessKey_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = ""
					If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then Response.Cookies("SIAP_Password_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = ""
					Session.Contents("SIAP_CurrentAccessKey") = ""
					Session.Contents("SIAP_CurrentPassword") = ""
				ElseIf (Not B_PORTAL) And (StrComp(aLoginComponent(S_PASSWORD_LOGIN), CStr(oRecordset.Fields("UserPassword").Value), vbBinaryCompare) <> 0) Then
					Call CheckLoginFailures(oADODBConnection, aLoginComponent, True)
					lErrorNumber = L_ERR_INCORRECT_PASSWORD
					sErrorDescription = "La clave de acceso y/o la contraseña <FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>están incorrectos</B></FONT>. Favor de introducirlos de nuevo. Recuerde que esta información <FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>es sensible a mayúsculas y minúsculas</B></FONT>."
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "LoginComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
					If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then Response.Cookies("SIAP_AccessKey_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = ""
					If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) <> 0 Then Response.Cookies("SIAP_Password_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = ""
					Session.Contents("SIAP_CurrentAccessKey") = ""
					Session.Contents("SIAP_CurrentPassword") = ""
				Else
					If CInt(oRecordset.Fields("SecurityLock").Value) >= CInt(GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION)) Then
						Response.Redirect "Logout.asp?LockedUser=1"
					Else
						If Len(oRequest("Login").Item) > 0 Then
							sErrorDescription = "No se pudo modificar la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set SecurityLock=0 Where (UserAccessKey='" & Replace(aLoginComponent(S_ACCESS_KEY_LOGIN), "'", "") & "')", "LoginComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
					Call CheckLoginFailures(oADODBConnection, aLoginComponent, False)
					Session.Contents("SIAP_CurrentAccessKey") = aLoginComponent(S_ACCESS_KEY_LOGIN)
					If B_PORTAL Then
						Session.Contents("SIAP_CurrentPassword") = aLoginComponent(S_ACCESS_KEY_LOGIN)
					Else
						Session.Contents("SIAP_CurrentPassword") = aLoginComponent(S_PASSWORD_LOGIN)
					End If
					aLoginComponent(B_VALID_USER_LOGIN) = True
					aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value)
					aLoginComponent(S_USER_NAME_LOGIN) = CStr(oRecordset.Fields("UserName").Value)
					aLoginComponent(S_USER_LAST_NAME_LOGIN) = CStr(oRecordset.Fields("UserLastName").Value)
					aLoginComponent(N_PROFILE_ID_LOGIN) = CLng(oRecordset.Fields("ProfileID").Value)
					aLoginComponent(N_USER_PERMISSIONS_LOGIN) = CStr(oRecordset.Fields("UserPermissions").Value)
					aLoginComponent(N_USER_PERMISSIONS2_LOGIN) = CStr(oRecordset.Fields("UserPermissions2").Value)
					aLoginComponent(N_USER_PERMISSIONS3_LOGIN) = CStr(oRecordset.Fields("UserPermissions3").Value)
					aLoginComponent(N_USER_PERMISSIONS4_LOGIN) = CStr(oRecordset.Fields("UserPermissions4").Value)
					aLoginComponent(L_PERMISSION_REPORTS_LOGIN) = CLng(oRecordset.Fields("PermissionReports").Value)
					aLoginComponent(L_PERMISSION_REPORTS2_LOGIN) = CLng(oRecordset.Fields("PermissionReports2").Value)
					aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) = CStr(oRecordset.Fields("PermissionAreaID").Value)
					aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) = CLng(oRecordset.Fields("PermissionZoneID").Value)
					aLoginComponent(S_PERMISSION_AREA_PATH_LOGIN) = ",-1,"
					aLoginComponent(S_PERMISSION_ZONE_PATH_LOGIN) = CStr(oRecordset.Fields("ZonePath").Value)
					aLoginComponent(S_USER_E_MAIL_LOGIN) = CStr(oRecordset.Fields("UserEmail").Value)
					aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN) = CStr(oRecordset.Fields("AdditionalEmail").Value)
					aLoginComponent(S_EMPLOYEE_NUMBER_LOGIN) = CStr(oRecordset.Fields("AdditionalEmail").Value)
					aLoginComponent(B_ACTIVE_LOGIN) = (CInt(oRecordset.Fields("UserActive").Value) <> 0)
					aLoginComponent(B_EXPIRED_LOGIN) = DateDiff("d", DateSerial(Left(CStr(oRecordset.Fields("ChangeDate").Value), Len("0000")), Mid(CStr(oRecordset.Fields("ChangeDate").Value), 5, Len("00")), Right(CStr(oRecordset.Fields("ChangeDate").Value), Len("00"))), Now()) >= CInt(GetAdminOption(aAdminOptionsComponent, PASSWORDS_DAYS_OPTION))
					aLoginComponent(B_BLOCKED_LOGIN) = (CInt(oRecordset.Fields("UserBlocked").Value) <> 0)
					aLoginComponent(B_TECH_SUPPORT_LOGIN) = (CInt(oRecordset.Fields("TechSupport").Value) <> 0)
					oRecordset.Close

					If Len(oRequest("Login").Item) > 0 Then
						sErrorDescription = "No se pudo registrar la entrada del usuario al sistema."
						Call ExecuteInsertQuerySp(oADODBConnection, "Insert Into SystemLogs (UserID, LogYear, LogMonth, LogDay, LogWeekDay, LogDate, LogHour, LogMinute, LogSecond, IPAddress) Values (" & aLoginComponent(N_USER_ID_LOGIN) & ", " & Year(Date()) & ", " & Month(Date()) & ", " & Day(Date()) & ", " & Weekday(Date()) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & Hour(Time()) & ", " & Minute(Time()) & ", " & Second(Time()) & ", '" & Request.ServerVariables("REMOTE_ADDR") & "')", "LoginComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetLoginCredencials = lErrorNumber
	Err.Clear
End Function

Function RedirectUser(oRequest, oADODBConnection, aLoginComponent, aOptionsComponent, lErrorNumber, sErrorDescription)
'************************************************************
'Purpose: To clean the cookies associated with the current
'         exam so the user can move to another course.
'Inputs:  oRequest, oADODBConnection, aLoginComponent, aOptionsComponent, lErrorNumber
'Outputs: lErrorNumber, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RedirectUser"
	Dim sStartPage
	Dim sRedirectURL

'	If DateDiff("d", Now(), DateSerial(2005, 04, 20)) < 0 Then
'		Response.Redirect "Default.asp?ExpiredLicense=1&DisplayError=1"
'	ElseIf (InStr(1, ("," & SERVER_NAME_FOR_LICENSE & "," & SERVER_IP_FOR_LICENSE & "," & EXT_SERVER_IP_FOR_LICENSE & ","), Request.ServerVariables("SERVER_NAME"), vbTextCompare) = 0) Then
'		Response.Redirect "Default.asp?InvalidLicense=1&DisplayError=1"
'	End If

'*** DO LOGIN ***'
	sRedirectURL = ""
	sStartPage = GetOption(aOptionsComponent, START_PAGE_OPTION)
	If InStr(1, sStartPage, "?", vbBinaryCompare) > 0 Then
		sStartPage = sStartPage & "&"
	Else
		sStartPage = sStartPage & "?"
	End If

	If (Len(oRequest("Login").Item) > 0) Then
		'The user has logged in. The list of logged users must be updated.
		If (aLoginComponent(B_VALID_USER_LOGIN)) Then
			If Len(oRequest("Redirect").Item) > 0 Then
				If Len(sRedirectURL) = 0 Then sRedirectURL = oRequest("Redirect").Item
			End If
			If Len(sRedirectURL) = 0 Then sRedirectURL = sStartPage & "Session=" & GetSerialNumberForDate("")
		End If
	End If

	If ((lErrorNumber <> 0) Or (Not aLoginComponent(B_VALID_USER_LOGIN))) And (StrComp(GetASPFileName(""), "Default.asp", vbTextCompare) <> 0) And (Len(Request.Cookies("SIAP_CurrentAccessKey")) > 0) Then
		'The user is not logged or the connection was lost.
		If Len(sRedirectURL) = 0 Then sRedirectURL = "Default.asp?InvalidUser=1"
	ElseIf aLoginComponent(B_BLOCKED_LOGIN) And (InStr(1, ",CourseWarning.asp,SaveUserOption.asp,", GetASPFileName(""), vbTextCompare) = 0) Then
		If Len(sRedirectURL) = 0 Then sRedirectURL = "CourseWarning.asp"
	ElseIf Not aLoginComponent(B_BLOCKED_LOGIN) And aLoginComponent(B_EXPIRED_LOGIN) And (InStr(1, ",ChangePassword.asp,SaveUserOption.asp,", GetASPFileName(""), vbTextCompare) = 0) Then
		If Len(sRedirectURL) = 0 Then sRedirectURL = "ChangePassword.asp?Expired=1"
	ElseIf Len(aLoginComponent(S_ACCESS_KEY_LOGIN)) > 0 Then
		If aLoginComponent(B_VALID_USER_LOGIN) And (StrComp(GetASPFileName(""), "Default.asp", vbTextCompare) = 0) Then
			If Len(oRequest("Redirect").Item) > 0 Then
				If Len(sRedirectURL) = 0 Then sRedirectURL = oRequest("Redirect").Item
			End If
			If Len(sRedirectURL) = 0 Then sRedirectURL = sStartPage & "Session=" & GetSerialNumberForDate("")
		End If
	ElseIf aLoginComponent(B_VALID_USER_LOGIN) Then
		If Len(oRequest("Redirect").Item) > 0 Then
			If Len(sRedirectURL) = 0 Then sRedirectURL = oRequest("Redirect").Item
		End If
		If Len(sRedirectURL) = 0 Then sRedirectURL = sStartPage & "Session=" & GetSerialNumberForDate("")
	ElseIf (Not aLoginComponent(B_VALID_USER_LOGIN)) And (StrComp(GetASPFileName(""), "Default.asp", vbTextCompare) <> 0) Then
		If Len(sRedirectURL) = 0 Then sRedirectURL = "Default.asp?RedirectUser=1&Redirect=" & Server.URLEncode(Replace(Request.ServerVariables("PATH_INFO"), Replace(SYSTEM_PATH, S_HTTP & Request.ServerVariables("SERVER_NAME"), "", 1, 1, vbBinaryCompare), "", 1, 1, vbBinaryCompare) & "?" & CStr(oRequest))
	End If
	If ((Not FileExists(Server.MapPath("Startup.htm"), sErrorDescription)) And (Not FileExists(Server.MapPath("Startup_" & aLoginComponent(N_PROFILE_ID_LOGIN) & ".htm"), sErrorDescription))) Or (StrComp(GetASPFileName(""), "Default.asp", vbBinaryCompare) <> 0) Or (Len(oRequest) = 0) Then
		If Len(sRedirectURL) > 0 Then Response.Redirect sRedirectURL
	ElseIf (Len(oRequest("Login").Item) > 0) Then
		If aLoginComponent(B_BLOCKED_LOGIN) Then sRedirectURL = "CourseWarning.asp"
		If aLoginComponent(B_EXPIRED_LOGIN) Then sRedirectURL = "ChangePassword.asp?Expired=1"
		sAuxMessage = ""
		If FileExists(Server.MapPath("Startup.htm"), sErrorDescription) Then
			sAuxMessage = sAuxMessage & GetFileContents(Server.MapPath("Startup.htm"), sErrorDescription)
		End If
		If FileExists(Server.MapPath("Startup_" & aLoginComponent(N_PROFILE_ID_LOGIN) & ".htm"), sErrorDescription) Then
			If Len(sAuxMessage) > 0 Then sAuxMessage = sAuxMessage & "<HR />"
			sAuxMessage = sAuxMessage & GetFileContents(Server.MapPath("Startup_" & aLoginComponent(N_PROFILE_ID_LOGIN) & ".htm"), sErrorDescription)
		End If
		If FileExists(Server.MapPath("Startup_" & Right(("000000" & aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN)), Len("000000")) & ".htm"), sErrorDescription) Then
			If Len(sAuxMessage) > 0 Then sAuxMessage = sAuxMessage & "<HR />"
			sAuxMessage = sAuxMessage & GetFileContents(Server.MapPath("Startup_" & Right(("000000" & aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN)), Len("000000")) & ".htm"), sErrorDescription)
		End If
		sAuxMessage = Replace(sAuxMessage, "<ACCESS_KEY />", aLoginComponent(S_ACCESS_KEY_LOGIN))
		sAuxMessage = Replace(sAuxMessage, "<PASSWORD />", aLoginComponent(S_PASSWORD_LOGIN))
		sAuxMessage = sAuxMessage & "<BR /><BR /><FONT FACE=""Arial"" SIZE=""3""><B><A HREF=""" & sRedirectURL & """>Continuar</A></B></FONT>"
	End If
End Function

Function CheckLoginFailures(oADODBConnection, aLoginComponent, bFails)
'************************************************************
'Purpose: To check if the login failures have exceeded the
'         admin option and to block the system in such case
'Inputs:  oADODBConnection, aLoginComponent, bFails
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckLoginFailures"
	Dim sIPAddress
	Dim iAttempts
	Dim oRecordset
	Dim lErrorNumber

	sIPAddress = Replace(Request.ServerVariables("REMOTE_ADDR"), ".", "_", 1, -1, vbBinaryCompare)
	If bFails Then
		If bCheckUserLoginOnly Then
			iAttempts = 0
			sErrorDescription = "No se pudieron obtener las credenciales de entrada del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SecurityLock From Users Where (UserAccessKey='" & Replace(aLoginComponent(S_ACCESS_KEY_LOGIN), "'", "") & "')", "LoginComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					iAttempts = CInt(oRecordset.Fields("SecurityLock").Value) + 1
					oRecordset.Close
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set SecurityLock=" & iAttempts & " Where (UserAccessKey='" & Replace(aLoginComponent(S_ACCESS_KEY_LOGIN), "'", "") & "')", "LoginComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If iAttempts >= CInt(GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION)) Then
						Response.Redirect "Logout.asp?LockedUser=1"
					End If
				End If
			End If
		Else
			If Len(Application.Contents("SIAP_" & sIPAddress)) > 0 Then
				Application.Contents("SIAP_" & sIPAddress) = CInt(Application.Contents("SIAP_" & sIPAddress)) + 1
			Else
				Application.Contents("SIAP_" & sIPAddress) = 1
			End If
			If CInt(Application.Contents("SIAP_" & sIPAddress)) >= CInt(GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION)) Then
				Application.Contents("SIAP_" & sIPAddress) = 0
				Call BlockTheSystem()
				Response.Redirect "Blocked.asp"
			End If
		End If
	Else
		Application.Contents("SIAP_" & sIPAddress) = Empty
	End If

	Set oRecordset = Nothing
	CheckLoginFailures = lErrorNumber
	Err.Clear
End Function
%>