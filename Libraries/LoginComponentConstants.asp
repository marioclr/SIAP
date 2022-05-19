<%
Const S_ACCESS_KEY_LOGIN = 0
Const S_PASSWORD_LOGIN = 1
Const B_VALID_USER_LOGIN = 2
Const B_EXPIRED_LOGIN = 3
Const B_BLOCKED_LOGIN = 4
Const N_USER_ID_LOGIN = 5
Const S_USER_NAME_LOGIN = 6
Const S_USER_LAST_NAME_LOGIN = 7
Const N_PROFILE_ID_LOGIN = 8
Const N_USER_PERMISSIONS_LOGIN = 9
Const N_USER_PERMISSIONS2_LOGIN = 10
Const N_USER_PERMISSIONS3_LOGIN = 11
Const N_USER_PERMISSIONS4_LOGIN = 12
Const L_PERMISSION_REPORTS_LOGIN = 13
Const L_PERMISSION_REPORTS2_LOGIN = 14
Const N_PERMISSION_AREA_ID_LOGIN = 15
Const N_PERMISSION_ZONE_ID_LOGIN = 16
Const S_PERMISSION_AREA_PATH_LOGIN = 17
Const S_PERMISSION_ZONE_PATH_LOGIN = 18
Const S_USER_E_MAIL_LOGIN = 19
Const S_USER_ADDITIONAL_E_MAIL_LOGIN = 20
Const S_EMPLOYEE_NUMBER_LOGIN = 21
Const B_ACTIVE_LOGIN = 22
Const B_TECH_SUPPORT_LOGIN = 23
Const S_PATH_LOGIN = 24
Const B_COMPONENT_INITIALIZED_LOGIN = 25

Const N_LOGIN_COMPONENT_SIZE = 25

Dim aLoginComponent()
Redim aLoginComponent(N_LOGIN_COMPONENT_SIZE)

Call InitializePermissionsForLoginComponent(oRequest, aLoginComponent)

Function InitializePermissionsForLoginComponent(oRequest, aLoginComponent)
'************************************************************
'Purpose: To initialize the permissions elements of the Login
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aLoginComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePermissionsForLoginComponent"
	Dim oRecordset
	Dim lErrorNumber
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

	If Len(aLoginComponent(S_ACCESS_KEY_LOGIN)) > 0 Then
		aLoginComponent(N_USER_ID_LOGIN) = -2
		sErrorDescription = "No se pudieron obtener las credenciales de entrada del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Users.*, ZonePath From Users, Zones Where (Users.PermissionZoneID=Zones.ZoneID) And (UserAccessKey='" & Replace(aLoginComponent(S_ACCESS_KEY_LOGIN), "'", "") & "')", "LoginComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value)
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
			Else
				aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value)
				aLoginComponent(N_PROFILE_ID_LOGIN) = -1
				aLoginComponent(N_USER_PERMISSIONS_LOGIN) = 0
				aLoginComponent(N_USER_PERMISSIONS2_LOGIN) = 0
				aLoginComponent(N_USER_PERMISSIONS3_LOGIN) = 0
				aLoginComponent(N_USER_PERMISSIONS4_LOGIN) = 0
				aLoginComponent(L_PERMISSION_REPORTS_LOGIN) = 0
				aLoginComponent(L_PERMISSION_REPORTS2_LOGIN) = 0
				aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) = "-2"
				aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) = -1
				aLoginComponent(S_PERMISSION_AREA_PATH_LOGIN) = ",-2,"
				aLoginComponent(S_PERMISSION_ZONE_PATH_LOGIN) = ",-2,"
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	InitializePermissionsForLoginComponent = lErrorNumber
	Err.Clear
End Function
%>