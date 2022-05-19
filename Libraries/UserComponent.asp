<%
Const N_ID_USER = 0
Const S_ACCESS_KEY_USER = 1
Const S_PASSWORD_USER = 2
Const S_NAME_USER = 3
Const S_LAST_NAME_USER = 4
Const S_EMAIL_USER = 5
Const L_PERMISSIONS_USER = 6
Const L_PERMISSIONS2_USER = 7
Const L_PERMISSIONS3_USER = 8
Const L_PERMISSIONS4_USER = 9
Const L_PERMISSION_REPORTS_USER = 10
Const L_PERMISSION_REPORTS2_USER = 11
Const S_PERMISSIONS_AREAS_USER = 12
Const L_PERMISSIONS_ZONE_USER = 13
Const S_PERMISSIONS_AREA_PATH_USER = 14
Const S_BOSS_EMAIL_USER = 15
Const S_ADDITIONAL_EMAIL_USER = 16
Const N_PROFILE_ID_USER = 17
Const N_ACTIVE_USER = 18
Const N_BLOCKED_USER = 19
Const N_TECH_SUPPORT_USER = 20
Const S_OLD_PASSWORD_USER = 21
Const B_CHECK_FOR_DUPLICATED_USER = 22
Const B_IS_DUPLICATED_USER = 23
Const B_COMPONENT_INITIALIZED_USER = 24

Const N_USER_COMPONENT_SIZE = 25

Dim aUserComponent()
Redim aUserComponent(N_USER_COMPONENT_SIZE)

Function InitializeUserComponent(oRequest, aUserComponent)
'************************************************************
'Purpose: To initialize the empty elements of the User Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aUserComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeUserComponent"
	Dim oItem
	Redim Preserve aUserComponent(N_USER_COMPONENT_SIZE)

	If IsEmpty(aUserComponent(N_ID_USER)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aUserComponent(N_ID_USER) = CLng(oRequest("UserID").Item)
		Else
			aUserComponent(N_ID_USER) = -2
		End If
	End If

	If IsEmpty(aUserComponent(S_ACCESS_KEY_USER)) Then
		If Len(oRequest("UserAccessKey").Item) > 0 Then
			aUserComponent(S_ACCESS_KEY_USER) = oRequest("UserAccessKey").Item
		Else
			aUserComponent(S_ACCESS_KEY_USER) = ""
		End If
	End If
	aUserComponent(S_ACCESS_KEY_USER) = Left(aUserComponent(S_ACCESS_KEY_USER), 120)

	If B_PORTAL Then
		aUserComponent(S_PASSWORD_USER) = aUserComponent(S_ACCESS_KEY_USER)
	Else
		If IsEmpty(aUserComponent(S_PASSWORD_USER)) Then
			If Len(oRequest("UserPassword").Item) > 0 Then
				aUserComponent(S_PASSWORD_USER) = oRequest("UserPassword").Item
			Else
				aUserComponent(S_PASSWORD_USER) = ""
			End If
		End If
		aUserComponent(S_PASSWORD_USER) = Left(aUserComponent(S_PASSWORD_USER), 120)
	End If

	If IsEmpty(aUserComponent(S_NAME_USER)) Then
		If Len(oRequest("UserName").Item) > 0 Then
			aUserComponent(S_NAME_USER) = oRequest("UserName").Item
		Else
			aUserComponent(S_NAME_USER) = ""
		End If
	End If
	aUserComponent(S_NAME_USER) = Left(aUserComponent(S_NAME_USER), 100)

	If IsEmpty(aUserComponent(S_LAST_NAME_USER)) Then
		If Len(oRequest("UserLastName").Item) > 0 Then
			aUserComponent(S_LAST_NAME_USER) = oRequest("UserLastName").Item
		Else
			aUserComponent(S_LAST_NAME_USER) = ""
		End If
	End If
	aUserComponent(S_LAST_NAME_USER) = Left(aUserComponent(S_LAST_NAME_USER), 100)

	If IsEmpty(aUserComponent(S_EMAIL_USER)) Then
		If Len(oRequest("UserEmail").Item) > 0 Then
			aUserComponent(S_EMAIL_USER) = oRequest("UserEmail").Item
		Else
			aUserComponent(S_EMAIL_USER) = ""
		End If
	End If
	aUserComponent(S_EMAIL_USER) = Left(aUserComponent(S_EMAIL_USER), 100)

	If IsEmpty(aUserComponent(L_PERMISSIONS_USER)) Then
		If Len(oRequest("UserPermissions").Item) > 0 Then
			If InStr(1, oRequest("UserPermissions").Item, ",", vbBinaryCompare) > 1 Then
				aUserComponent(L_PERMISSIONS_USER) = 0
				For Each oItem In oRequest("UserPermissions")
					aUserComponent(L_PERMISSIONS_USER) = aUserComponent(L_PERMISSIONS_USER) + CLng(oItem)
				Next
			Else
				aUserComponent(L_PERMISSIONS_USER) = CStr(oRequest("UserPermissions").Item)
			End If
		Else
			aUserComponent(L_PERMISSIONS_USER) = 0
		End If
	End If

	If IsEmpty(aUserComponent(L_PERMISSIONS4_USER)) Then
		If Len(oRequest("UserPermissions4").Item) > 0 Then
			aUserComponent(L_PERMISSIONS4_USER) = CStr(oRequest("UserPermissions4").Item)
		Else
			aUserComponent(L_PERMISSIONS4_USER) = 0
		End If
	End If

	If IsEmpty(aUserComponent(L_PERMISSION_REPORTS_USER)) Then
		If Len(oRequest("PermissionReports").Item) > 0 Then
			If InStr(1, oRequest("PermissionReports").Item, ",", vbBinaryCompare) > 1 Then
				aUserComponent(L_PERMISSION_REPORTS_USER) = 0
				For Each oItem In oRequest("PermissionReports")
					aUserComponent(L_PERMISSION_REPORTS_USER) = aUserComponent(L_PERMISSION_REPORTS_USER) + CLng(oItem)
				Next
			Else
				aUserComponent(L_PERMISSION_REPORTS_USER) = CLng(oRequest("PermissionReports").Item)
			End If
		Else
			aUserComponent(L_PERMISSION_REPORTS_USER) = 0
		End If
	End If

	If IsEmpty(aUserComponent(L_PERMISSION_REPORTS2_USER)) Then
		If Len(oRequest("PermissionReports2").Item) > 0 Then
			If InStr(1, oRequest("PermissionReports2").Item, ",", vbBinaryCompare) > 1 Then
				aUserComponent(L_PERMISSION_REPORTS2_USER) = 0
				For Each oItem In oRequest("PermissionReports2")
					aUserComponent(L_PERMISSION_REPORTS2_USER) = aUserComponent(L_PERMISSION_REPORTS2_USER) + CLng(oItem)
				Next
			Else
				aUserComponent(L_PERMISSION_REPORTS2_USER) = CLng(oRequest("PermissionReports2").Item)
			End If
		Else
			aUserComponent(L_PERMISSION_REPORTS2_USER) = 0
		End If
	End If

	If IsEmpty(aUserComponent(S_PERMISSIONS_AREAS_USER)) Then
		aUserComponent(S_PERMISSIONS_AREAS_USER) = "-2"
		If Len(oRequest("AreaID1").Item) > 0 Then
			aUserComponent(S_PERMISSIONS_AREAS_USER) = CLng(oRequest("AreaID1").Item)
		Else
			For Each oItem In oRequest("AreaID")
				aUserComponent(S_PERMISSIONS_AREAS_USER) = aUserComponent(S_PERMISSIONS_AREAS_USER) & "," & oItem
			Next
		End If
	End If

	If IsEmpty(aUserComponent(L_PERMISSIONS_ZONE_USER)) Then
		If Len(oRequest("PermissionZoneID").Item) > 0 Then
			aUserComponent(L_PERMISSIONS_ZONE_USER) = CLng(oRequest("PermissionZoneID").Item)
		Else
			aUserComponent(L_PERMISSIONS_ZONE_USER) = -1
		End If
	End If

	aUserComponent(S_PERMISSIONS_AREA_PATH_USER) = ""

	If IsEmpty(aUserComponent(S_BOSS_EMAIL_USER)) Then
		If Len(oRequest("BossEmail").Item) > 0 Then
			aUserComponent(S_BOSS_EMAIL_USER) = oRequest("BossEmail").Item
		Else
			aUserComponent(S_BOSS_EMAIL_USER) = ""
		End If
	End If
	aUserComponent(S_BOSS_EMAIL_USER) = Left(aUserComponent(S_BOSS_EMAIL_USER), 100)

	If IsEmpty(aUserComponent(S_ADDITIONAL_EMAIL_USER)) Then
		If Len(oRequest("AdditionalEmail").Item) > 0 Then
			aUserComponent(S_ADDITIONAL_EMAIL_USER) = oRequest("AdditionalEmail").Item
		Else
			aUserComponent(S_ADDITIONAL_EMAIL_USER) = ""
		End If
	End If
	aUserComponent(S_ADDITIONAL_EMAIL_USER) = Left(aUserComponent(S_ADDITIONAL_EMAIL_USER), 100)

	If IsEmpty(aUserComponent(N_PROFILE_ID_USER)) Then
		If Len(oRequest("ProfileID").Item) > 0 Then
			aUserComponent(N_PROFILE_ID_USER) = CLng(oRequest("ProfileID").Item)
		Else
			aUserComponent(N_PROFILE_ID_USER) = -1
		End If
	End If

	If IsEmpty(aUserComponent(L_PERMISSIONS2_USER)) Then
		Select Case aUserComponent(N_PROFILE_ID_USER)
			Case 1
				If Len(oRequest("UserPermissionsProfile1").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile1")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 2
				If Len(oRequest("UserPermissionsProfile2").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile2")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 3
				If Len(oRequest("UserPermissionsProfile3").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile3")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 4
				If Len(oRequest("UserPermissionsProfile4").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile4")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 5
				If Len(oRequest("UserPermissionsProfile5").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile5")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 6
				If Len(oRequest("UserPermissionsProfile6").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile6")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 7
				If Len(oRequest("UserPermissionsProfile7").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile7")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case 8
				If Len(oRequest("UserPermissionsProfile8").Item) > 0 Then
					For Each oItem In oRequest("UserPermissionsProfile8")
						aUserComponent(L_PERMISSIONS2_USER) = aUserComponent(L_PERMISSIONS2_USER) & "," & oItem
					Next
				Else
					aUserComponent(L_PERMISSIONS2_USER) = 0
				End If
			Case Else
				aUserComponent(L_PERMISSIONS2_USER) = "-1"
		End Select
	End If

'	If IsEmpty(aUserComponent(L_PERMISSIONS3_USER)) Then
		Select Case aUserComponent(N_PROFILE_ID_USER)
			Case 2
				If Len(oRequest("UserPermissions3b").Item) > 0 Then
					If InStr(1, oRequest("UserPermissions3b").Item, ",", vbBinaryCompare) > 1 Then
						aUserComponent(L_PERMISSIONS3_USER) = 0
						For Each oItem In oRequest("UserPermissions3b")
							aUserComponent(L_PERMISSIONS3_USER) = aUserComponent(L_PERMISSIONS3_USER) + CLng(oItem)
						Next
					Else
						aUserComponent(L_PERMISSIONS3_USER) = CLng(oRequest("UserPermissions3b").Item)
					End If
				Else
					aUserComponent(L_PERMISSIONS3_USER) = 0
				End If
			Case 7
				If Len(oRequest("UserPermissions3g").Item) > 0 Then
					If InStr(1, oRequest("UserPermissions3g").Item, ",", vbBinaryCompare) > 1 Then
						aUserComponent(L_PERMISSIONS3_USER) = 0
						For Each oItem In oRequest("UserPermissions3g")
							aUserComponent(L_PERMISSIONS3_USER) = aUserComponent(L_PERMISSIONS3_USER) + CLng(oItem)
						Next
					Else
						aUserComponent(L_PERMISSIONS3_USER) = CLng(oRequest("UserPermissions3g").Item)
					End If
				Else
					aUserComponent(L_PERMISSIONS3_USER) = 0
				End If
			Case Else
				If Len(oRequest("UserPermissions3").Item) > 0 Then
					If InStr(1, oRequest("UserPermissions3").Item, ",", vbBinaryCompare) > 1 Then
						aUserComponent(L_PERMISSIONS3_USER) = 0
						For Each oItem In oRequest("UserPermissions3")
							aUserComponent(L_PERMISSIONS3_USER) = aUserComponent(L_PERMISSIONS3_USER) + CLng(oItem)
						Next
					Else
						aUserComponent(L_PERMISSIONS3_USER) = CLng(oRequest("UserPermissions3").Item)
					End If
				Else
					aUserComponent(L_PERMISSIONS3_USER) = 0
				End If
		End Select
'	End If

	If IsEmpty(aUserComponent(N_ACTIVE_USER)) Then
		If Len(oRequest("UserActive").Item) > 0 Then
			aUserComponent(N_ACTIVE_USER) = CInt(oRequest("UserActive").Item)
		Else
			aUserComponent(N_ACTIVE_USER) = 0
		End If
	End If

	If IsEmpty(aUserComponent(N_BLOCKED_USER)) Then
		If Len(oRequest("UserBlocked").Item) > 0 Then
			aUserComponent(N_BLOCKED_USER) = CInt(oRequest("UserBlocked").Item)
		Else
			aUserComponent(N_BLOCKED_USER) = N_USER_BLOCKED
		End If
	End If

	If IsEmpty(aUserComponent(N_TECH_SUPPORT_USER)) Then
		If Len(oRequest("TechSupport").Item) > 0 Then
			aUserComponent(N_TECH_SUPPORT_USER) = CInt(oRequest("TechSupport").Item)
		Else
			aUserComponent(N_TECH_SUPPORT_USER) = 0
		End If
	End If

	aUserComponent(B_CHECK_FOR_DUPLICATED_USER) = True
	aUserComponent(B_IS_DUPLICATED_USER) = False

	aUserComponent(B_COMPONENT_INITIALIZED_USER) = True
	InitializeUserComponent = Err.number
	Err.Clear
End Function

Function AddUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new user into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddUser"
	Dim lNewUserID
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If aUserComponent(N_ID_USER) = -2 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo usuario."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Users", "UserID", "", 1, aUserComponent(N_ID_USER), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aUserComponent(B_CHECK_FOR_DUPLICATED_USER) Then
			lErrorNumber = CheckExistencyOfUser(oADODBConnection, False, aUserComponent, sErrorDescription)
			If aUserComponent(B_IS_DUPLICATED_USER) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "La clave de acceso " & aUserComponent(S_ACCESS_KEY_USER) & " o el empleado número " & aUserComponent(S_ADDITIONAL_EMAIL_USER) & " ya están registrados en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			ElseIf B_SADE Then
				If IsEmpty(oRequest("Import").Item) Then
					lErrorNumber = CheckExistencyOfUser(oADODBConnection, True, aUserComponent, sErrorDescription)
					If aUserComponent(B_IS_DUPLICATED_USER) Then
						lErrorNumber = L_ERR_DUPLICATED_RECORD
						sErrorDescription = "La clave de acceso " & aUserComponent(S_ACCESS_KEY_USER) & " ya está registrada en el sistema SADE."
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
					End If
				End If
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckUserInformationConsistency(aUserComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información del nuevo usuario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Users (UserID, UserAccessKey, UserPassword, UserName, UserLastName, UserEmail, UserPermissions, UserPermissions2, UserPermissions3, UserPermissions4, PermissionReports, PermissionReports2, PermissionAreaID, PermissionZoneID, BossEmail, AdditionalEmail, ProfileID, UserActive, UserBlocked, SecurityLock, TechSupport) Values (" & aUserComponent(N_ID_USER) & ", '" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "', '" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', '" & Replace(aUserComponent(S_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_EMAIL_USER), "'", "") & "', '" & aUserComponent(L_PERMISSIONS_USER) & "', '" & aUserComponent(L_PERMISSIONS2_USER) & "', '" & aUserComponent(L_PERMISSIONS3_USER) & "', '" & aUserComponent(L_PERMISSIONS4_USER) & "', " & aUserComponent(L_PERMISSION_REPORTS_USER) & ", " & aUserComponent(L_PERMISSION_REPORTS2_USER) & ", '" & aUserComponent(S_PERMISSIONS_AREAS_USER) & "', " & aUserComponent(L_PERMISSIONS_ZONE_USER) & ", '" & Replace(aUserComponent(S_BOSS_EMAIL_USER), "'", "") & "', '" & Replace(aUserComponent(S_ADDITIONAL_EMAIL_USER), "'", "") & "', " & aUserComponent(N_PROFILE_ID_USER) & ", " & aUserComponent(N_ACTIVE_USER) & ", " & aUserComponent(N_BLOCKED_USER) & ", 0, " & aUserComponent(N_TECH_SUPPORT_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo guardar la contraseña del nuevo usuario."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into UsersPWD (UserID, UserOldPassword, ChangeDate) Values (" & aUserComponent(N_ID_USER) & ", '" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If (lErrorNumber = 0) And (B_USE_SMTP) Then lErrorNumber = SendMessageToNewUser(aUserComponent, sErrorDescription)
				If (lErrorNumber = 0) And (B_SADE) Then
					lErrorNumber = GetNewIDFromTable(oSADEADODBConnection, "Usuario", "ID_Usuario", "", 1, lNewUserID, sErrorDescription)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información del nuevo usuario en la base de datos de SADE. Será necesario dar de alta los mismos datos directamente en SADE entrando al módulo de Administración de Grupos y Usuarios."
						lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Insert Into Usuario (ID_Usuario, ID_Grupo, ID_Tipo, ID_Empleado, Nombre, Apellidos, Clave_Acceso, Password_Acceso, e_mail, Curriculum, Permisos, ID_Privilegio, Competencias, Fecha_Ingreso, Fecha_Expiracion, Activo) Values (" & lNewUserID & ", " & SADE_GROUP_ID & ", 2, '" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "', '" & Replace(aUserComponent(S_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "', '" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', '" & Replace(aUserComponent(S_EMAIL_USER), "'", "") & "', '', 0, -1, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) + 10000 & ", " & aUserComponent(N_ACTIVE_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la descripción del nuevo usuario en la base de datos de SADE. Será necesario dar de alta los mismos datos directamente en SADE entrando al módulo de Administración de Grupos y Usuarios."
							lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Insert Into UsuarioDescripcion (ID_Usuario, Descripcion) Values (" & lNewUserID & ", 'La cuenta de este usuario fue dada de alta desde SICOSI')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron guardar los grupos del nuevo usuario en la base de datos de SADE. Será necesario dar de alta los mismos datos directamente en SADE entrando al módulo de Administración de Grupos y Usuarios."
								lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Insert Into GruposUsuariosLKP (ID_Usuario, ID_Grupo) Values (" & lNewUserID & ", " & SADE_GROUP_ID & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo guardar el perfil del nuevo usuario en la base de datos de SADE. Será necesario dar de alta los mismos datos directamente en SADE entrando al módulo de Administración de Grupos y Usuarios."
									lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Insert Into PerfilesUsuariosLKP (ID_Usuario, ID_Perfil) Values (" & lNewUserID & ", " & SADE_PROFILE_ID & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	AddUser = lErrorNumber
	Err.Clear
End Function

Function ImportUser(oRequest, oADODBConnection, oSADEADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To import an user from the SADE database
'Inputs:  oRequest, oADODBConnection, oSADEADODBConnection, 
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ImportUser"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If Len(aUserComponent(S_ACCESS_KEY_USER)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave de acceso del usuario para obtener su información desde la base de datos de SADE."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del usuario desde la base de datos de SADE."
		lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Select * From Usuario Where (Clave_Acceso='" & aUserComponent(S_ACCESS_KEY_USER) & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El usuario especificado no se encuentra en la base de datos de SADE."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aUserComponent(S_ACCESS_KEY_USER) = CStr(oRecordset.Fields("Clave_Acceso").Value)
				aUserComponent(S_PASSWORD_USER) = CStr(oRecordset.Fields("Password_Acceso").Value)
				aUserComponent(S_NAME_USER) = CStr(oRecordset.Fields("Nombre").Value)
				aUserComponent(S_LAST_NAME_USER) = CStr(oRecordset.Fields("Apellidos").Value)
				aUserComponent(S_EMAIL_USER) = " - - - "
				aUserComponent(S_BOSS_EMAIL_USER) = " - - - "
				If Not IsNull(oRecordset.Fields("e_mail").Value) Then
					If Len(CStr(oRecordset.Fields("e_mail").Value)) > 0 Then
						aUserComponent(S_EMAIL_USER) = CStr(oRecordset.Fields("e_mail").Value)
						aUserComponent(S_BOSS_EMAIL_USER) = CStr(oRecordset.Fields("e_mail").Value)
					End If
				End If
				aUserComponent(L_PERMISSIONS_USER) = 0
				aUserComponent(L_PERMISSIONS2_USER) = 0
				aUserComponent(L_PERMISSIONS3_USER) = 0
				aUserComponent(L_PERMISSIONS4_USER) = 0
				aUserComponent(L_PERMISSION_REPORTS_USER) = 0
				aUserComponent(L_PERMISSION_REPORTS2_USER) = 0
				aUserComponent(S_PERMISSIONS_AREAS_USER) = "-2"
				aUserComponent(L_PERMISSIONS_ZONE_USER) = -1
				aUserComponent(S_ADDITIONAL_EMAIL_USER) = ""
				aUserComponent(N_PROFILE_ID_USER) = -1
				aUserComponent(N_ACTIVE_USER) = CInt(oRecordset.Fields("Activo").Value)
				aUserComponent(N_BLOCKED_USER) = 0
				aUserComponent(N_TECH_SUPPORT_USER) = 0
				sErrorDescription = "No se pudo obtener un identificador para el nuevo usuario."
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "Users", "UserID", "", 1, aUserComponent(N_ID_USER), sErrorDescription)
				If Not CheckUserInformationConsistency(aUserComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo guardar la información del nuevo usuario."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Users (UserID, UserAccessKey, UserPassword, UserName, UserLastName, UserEmail, UserPermissions, UserPermissions2, UserPermissions3, UserPermissions4, PermissionReports, PermissionReports2, PermissionAreaID, PermissionZoneID, BossEmail, AdditionalEmail, ProfileID, UserActive, UserBlocked, SecurityLock, TechSupport) Values (" & aUserComponent(N_ID_USER) & ", '" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "', '" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', '" & Replace(aUserComponent(S_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "', '" & Replace(aUserComponent(S_EMAIL_USER), "'", "") & "', '" & aUserComponent(L_PERMISSIONS_USER) & "', '" & aUserComponent(L_PERMISSIONS2_USER) & "', '" & aUserComponent(L_PERMISSIONS3_USER) & "', '" & aUserComponent(L_PERMISSIONS4_USER) & "', " & aUserComponent(L_PERMISSION_REPORTS_USER) & ", " & aUserComponent(L_PERMISSION_REPORTS2_USER) & ", '" & aUserComponent(S_PERMISSIONS_AREAS_USER) & "', " & aUserComponent(L_PERMISSIONS_ZONE_USER) & ", '" & Replace(aUserComponent(S_BOSS_EMAIL_USER), "'", "") & "', '" & Replace(aUserComponent(S_ADDITIONAL_EMAIL_USER), "'", "") & "', " & aUserComponent(N_PROFILE_ID_USER) & ", " & aUserComponent(N_ACTIVE_USER) & ", " & aUserComponent(N_BLOCKED_USER) & ", 0, " & aUserComponent(N_TECH_SUPPORT_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la contraseña del nuevo usuario."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into UsersPWD (UserID, UserOldPassword, ChangeDate) Values (" & aUserComponent(N_ID_USER) & ", '" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	ImportUser = lErrorNumber
	Err.Clear
End Function

Function GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an user from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetUser"
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If (aUserComponent(N_ID_USER) = -2) And (StrComp(GetASPFileName(""), "SendPassword.asp", vbBinaryCompare) <> 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del usuario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sCondition = "(Users.UserID=" & aUserComponent(N_ID_USER) & ")"
		If (StrComp(GetASPFileName(""), "SendPassword.asp", vbBinaryCompare) = 0) And (Len(oRequest("UserAccessKey").Item) > 0) Then sCondition = "(Users.UserAccessKey='" & Replace(oRequest("UserAccessKey").Item, "'", "") & "')"
		sErrorDescription = "No se pudo obtener la información del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Users.*, UserOldPassword From Users, UsersPWD Where (Users.UserID=UsersPWD.UserID) And " & sCondition, "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El usuario especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aUserComponent(S_ACCESS_KEY_USER) = CStr(oRecordset.Fields("UserAccessKey").Value)
				aUserComponent(S_PASSWORD_USER) = CStr(oRecordset.Fields("UserPassword").Value)
				aUserComponent(S_NAME_USER) = CStr(oRecordset.Fields("UserName").Value)
				aUserComponent(S_LAST_NAME_USER) = CStr(oRecordset.Fields("UserLastName").Value)
				aUserComponent(S_EMAIL_USER) = CStr(oRecordset.Fields("UserEmail").Value)
				aUserComponent(L_PERMISSIONS_USER) = CStr(oRecordset.Fields("UserPermissions").Value)
				aUserComponent(L_PERMISSIONS2_USER) = CStr(oRecordset.Fields("UserPermissions2").Value)
				aUserComponent(L_PERMISSIONS3_USER) = CStr(oRecordset.Fields("UserPermissions3").Value)
				aUserComponent(L_PERMISSIONS4_USER) = CStr(oRecordset.Fields("UserPermissions4").Value)
				aUserComponent(L_PERMISSION_REPORTS_USER) = CLng(oRecordset.Fields("PermissionReports").Value)
				aUserComponent(L_PERMISSION_REPORTS2_USER) = CLng(oRecordset.Fields("PermissionReports2").Value)
				aUserComponent(S_PERMISSIONS_AREAS_USER) = CStr(oRecordset.Fields("PermissionAreaID").Value)
				aUserComponent(L_PERMISSIONS_ZONE_USER) = CLng(oRecordset.Fields("PermissionZoneID").Value)
				aUserComponent(S_PERMISSIONS_AREA_PATH_USER) = "-1"
				aUserComponent(S_BOSS_EMAIL_USER) = CStr(oRecordset.Fields("BossEmail").Value)
				aUserComponent(S_ADDITIONAL_EMAIL_USER) = CStr(oRecordset.Fields("AdditionalEmail").Value)
				aUserComponent(N_PROFILE_ID_USER) = CLng(oRecordset.Fields("ProfileID").Value)
				aUserComponent(N_ACTIVE_USER) = CInt(oRecordset.Fields("UserActive").Value)
				aUserComponent(N_BLOCKED_USER) = CInt(oRecordset.Fields("UserBlocked").Value)
				aUserComponent(N_TECH_SUPPORT_USER) = CInt(oRecordset.Fields("TechSupport").Value)
				aUserComponent(S_OLD_PASSWORD_USER) = CStr(oRecordset.Fields("UserOldPassword").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetUser = lErrorNumber
	Err.Clear
End Function

Function GetUsers(oRequest, oADODBConnection, aUserComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the users from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetUsers"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If
	sCondition = ""
	If Len(oRequest("StartFrom").Item) > 0 Then
		sCondition = " And (Users.UserID>=" & oRequest("StartFrom").Item & ")"
	End If

	If InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) <> 1 Then sCondition = sCondition & "And (UserID > 9)"
	sErrorDescription = "No se pudo obtener la información de los usuarios."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Users Where (UserID > -1) " & sCondition & " Order By UserLastName, UserName", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetUsers = lErrorNumber
	Err.Clear
End Function

Function ModifyUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing user in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyUser"
	Dim sCurrentPassword
	Dim sOldPassword
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

    lErrorNumber = CreateADODBConnection(SIAP_DATABASE_PATH, SIAP_DATABASE_USERNAME, SIAP_DATABASE_PASSWORD, iConnectionType, oADODBConnection, sErrorDescription)
    If(lErrorNumber=0) Then    
        bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	    If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		    Call InitializeUserComponent(oRequest, aUserComponent)
	    End If

	    If aUserComponent(N_ID_USER) = -2 Then
		    lErrorNumber = -1
		    sErrorDescription = "No se especificó el identificador del usuario a modificar."
		    Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	    Else
		    If Not CheckUserInformationConsistency(aUserComponent, sErrorDescription) Then
			    lErrorNumber = -1
		    Else
			    'If aUserComponent(N_PROFILE_ID_USER) = 4 Then
				'    aUserComponent(L_PERMISSIONS2_USER) = 0
				'    aUserComponent(L_PERMISSIONS3_USER) = 0
			    'End If
			    sErrorDescription = "No se pudo modificar la información del usuario."
			    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set UserName='" & Replace(aUserComponent(S_NAME_USER), "'", "") & "', UserLastName='" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "', UserEmail='" & Replace(aUserComponent(S_EMAIL_USER), "'", "") & "', UserPermissions='" & aUserComponent(L_PERMISSIONS_USER) & "', UserPermissions2='" & aUserComponent(L_PERMISSIONS2_USER) & "', UserPermissions3=" & aUserComponent(L_PERMISSIONS3_USER) & ", UserPermissions4=" & aUserComponent(L_PERMISSIONS4_USER) & ", PermissionReports=" & aUserComponent(L_PERMISSION_REPORTS_USER) & ", PermissionReports2=" & aUserComponent(L_PERMISSION_REPORTS2_USER) & ", PermissionAreaID='" & aUserComponent(S_PERMISSIONS_AREAS_USER) & "', PermissionZoneID=" & aUserComponent(L_PERMISSIONS_ZONE_USER) & ", BossEmail='" & Replace(aUserComponent(S_BOSS_EMAIL_USER), "'", "") & "', AdditionalEmail='" & Replace(aUserComponent(S_ADDITIONAL_EMAIL_USER), "'", "") & "', ProfileID=" & aUserComponent(N_PROFILE_ID_USER) & ", UserActive=" & aUserComponent(N_ACTIVE_USER) & ", UserBlocked=" & aUserComponent(N_BLOCKED_USER) & ", TechSupport=" & aUserComponent(N_TECH_SUPPORT_USER) & " Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			    If lErrorNumber = 0 Then
				    sErrorDescription = "No se pudo modificar la contraseña del usuario."
				    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserPassword From Users Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				    If lErrorNumber = 0 Then
					    sCurrentPassword = CStr(oRecordset.Fields("UserPassword").Value)
					    oRecordset.Close
					    If StrComp(sCurrentPassword, aUserComponent(S_PASSWORD_USER), vbBinaryCompare) <> 0 Then
						    sErrorDescription = "No se pudo modificar la contraseña del usuario."
						    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserOldPassword From UsersPWD Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						    If lErrorNumber = 0 Then
							    sOldPassword = CStr(oRecordset.Fields("UserOldPassword").Value)
							    oRecordset.Close
						    End If
						    If StrComp(aUserComponent(S_PASSWORD_USER), sOldPassword, vbBinaryCompare) = 0 Then
							    lErrorNumber = -1
							    sErrorDescription = "La contraseña especificada ya había sido utilizada con anterioridad. Por razones de seguridad, deberá introducir otra contraseña."
						    Else
							    sErrorDescription = "No se pudo modificar la información del usuario."
							    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set UserPassword='" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "' Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

							    sErrorDescription = "No se pudo modificar la contraseña del usuario."
							    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update UsersPWD Set UserOldPassword='" & sCurrentPassword & "', ChangeDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						    End If
					    End If
				    End If
			    End If
			    If lErrorNumber = 0 Then
				    If B_SADE Then
					    sErrorDescription = "No se pudo modificar la información del usuario en la base de datos de SADE. Si hubo cambios en el nombre o el apellido del usuario, en su contraseña o en su cuenta de correo electrónico, será necesario realizar el mismo cambio directamente en SADE entrando al módulo de Administración de Grupos y Usuarios."
					    lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Update Usuario Set Nombre='" & Replace(aUserComponent(S_NAME_USER), "'", "") & "', Apellidos='" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "', Password_Acceso='" & Replace(aUserComponent(S_PASSWORD_USER), "'", "") & "', e_mail='" & Replace(aUserComponent(S_EMAIL_USER), "'", "") & "' Where (Clave_Acceso='" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				    End If
				    If lErrorNumber = 0 Then
					    If aLoginComponent(N_USER_ID_LOGIN) = aUserComponent(N_ID_USER) Then
						    Response.Cookies("SIAP_CurrentPassword") = aUserComponent(S_PASSWORD_USER)
						    Response.Cookies("SIAP_Password_" & aLoginComponent(S_ACCESS_KEY_LOGIN)) = aUserComponent(S_PASSWORD_USER)
						    If B_SADE Then
							    Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewline
								    Response.Write "alert('Si usted se encuentra utilizando SADE le recomendamos que cierre\n su sesión presionando SALIR y vuelva a entrar al sistema.');" & vbNewline
							    Response.Write "//--></SCRIPT>" & vbNewline
						    End If
					    End If
				    End If
			    End If
		    End If
	    End If
    End If

	

	Set oRecordset = Nothing
	ModifyUser = lErrorNumber
	Err.Clear
End Function

Function SetActiveForUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given user
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForUser"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If aUserComponent(N_ID_USER) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set UserActive=" & CInt(oRequest("SetActive").Item) & " Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForUser = lErrorNumber
	Err.Clear
End Function

Function RemoveUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an user from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveUser"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If aUserComponent(N_ID_USER) = -2 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el usuario a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Users Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la contraseña del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From UsersPWD Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron eliminar las preferencias del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Preferences Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la contraseña del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From UsersReportsLKP Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron eliminar las entradas al sistema del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From SystemLogs Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasHistoryList Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasPositionsHistoryList Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasPositionsHistoryList Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptStateTaxLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptStateTaxLKP Set EndUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set AddUserID=-1 Where (AddUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set RemoveUserID=-1 Where (RemoveUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHandicapsLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHandicapsLKP Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesInformation Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRequirementsLKP Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSyndicatesLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSyndicatesLKP Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsBudgetsLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsBudgetsLKP Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsConceptsLKP Set StartUserID=-1 Where (StartUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsConceptsLKP Set EndUserID=-1 Where (EndUserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsLevelsHistoryList Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsRequirementsLKP Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron actualizar los reportes del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Reports Set UserID=-1 Where (UserID=" & aUserComponent(N_ID_USER) & ")", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveUser = lErrorNumber
	Err.Clear
End Function

Function SendMessageToNewUser(aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To send a message to alert the new user about
'         his/her account in Workflow
'Inputs:  aUserComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SendMessageToNewUser"
	Dim sTemplate
	Dim lErrorNumber

	If B_USE_SMTP And (Not B_PORTAL) And (Len(aUserComponent(S_EMAIL_USER)) > 0) Then
		sTemplate = ""
		If Len(oRequest("TemplateFile").Item) = 0 Then
			sTemplate = GetFileContents(Server.MapPath("Template_NewUser.htm"), sErrorDescription)
		Else
			If FileExists(Server.MapPath(oRequest("TemplateFile").Item), sErrorDescription) Then
				sTemplate = GetFileContents(Server.MapPath(oRequest("TemplateFile").Item), sErrorDescription)
			Else
				sTemplate = GetFileContents(Server.MapPath("Template_NewUser.htm"), sErrorDescription)
			End If
		End If
		sTemplate = Replace(sTemplate, "<USER_NAME />", aUserComponent(S_NAME_USER) & " " & aUserComponent(S_LAST_NAME_USER))
		sTemplate = Replace(sTemplate, "<ACCESS_KEY />", aUserComponent(S_ACCESS_KEY_USER))
		sTemplate = Replace(sTemplate, "<PASSWORD />", aUserComponent(S_PASSWORD_USER))
		sTemplate = Replace(sTemplate, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		If Len(sTemplate) > 0 Then
			ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
			aEmailComponent(S_TO_EMAIL) = aUserComponent(S_EMAIL_USER)
			aEmailComponent(S_CC_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_SUBJECT_EMAIL) = "Activación de cuenta en el Sistema de Administración del Personal"
			aEmailComponent(S_BODY_EMAIL) = sTemplate
			lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
		End If
	End If

	SendMessageToNewUser = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfUser(oADODBConnection, bInSADE, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific user exists in the database
'Inputs:  oADODBConnection, bInSADE, aUserComponent
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfUser"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aUserComponent(B_COMPONENT_INITIALIZED_USER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeUserComponent(oRequest, aUserComponent)
	End If

	If (Len(aUserComponent(S_ACCESS_KEY_USER)) = 0) And ((Len(aUserComponent(S_NAME_USER)) = 0) Or (Len(aUserComponent(S_LAST_NAME_USER)) = 0)) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave de acceso o el nombre y el apellido del usuario para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	If Len(aUserComponent(S_ACCESS_KEY_USER)) > 0 Then
		sErrorDescription = "No se pudo revisar la existencia del usuario en la base de datos."
		If bInSADE Then
			lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Select * From Usuario Where (Clave_Acceso='" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Else
			'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Users Where (UserAccessKey='" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "') Or (AdditionalEmail='" & Replace(aUserComponent(S_ADDITIONAL_EMAIL_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Users Where (UserAccessKey='" & Replace(aUserComponent(S_ACCESS_KEY_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		If lErrorNumber = 0 Then
			aUserComponent(B_IS_DUPLICATED_USER) = (Not oRecordset.EOF)
			oRecordset.Close
		End If
	ElseIf (Len(aUserComponent(S_NAME_USER)) > 0) And (Len(aUserComponent(S_LAST_NAME_USER)) > 0) Then
		sErrorDescription = "No se pudo revisar la existencia del usuario en la base de datos."
		If bInSADE Then
			lErrorNumber = ExecuteSQLQuery(oSADEADODBConnection, "Select * From Usuario Where (Nombre='" & Replace(aUserComponent(S_NAME_USER), "'", "") & "') And (Apellidos='" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Users Where (UserName='" & Replace(aUserComponent(S_NAME_USER), "'", "") & "') And (UserLastName='" & Replace(aUserComponent(S_LAST_NAME_USER), "'", "") & "')", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		If lErrorNumber = 0 Then
			aUserComponent(B_IS_DUPLICATED_USER) = (Not oRecordset.EOF)
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	CheckExistencyOfUser = lErrorNumber
	Err.Clear
End Function

Function CheckUserInformationConsistency(aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aUserComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckUserInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aUserComponent(N_ID_USER)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del usuario no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aUserComponent(S_ACCESS_KEY_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave de acceso del usuario está vacía."
		bIsCorrect = False
	End If
	If B_PORTAL Then aUserComponent(S_PASSWORD_USER) = aUserComponent(S_ACCESS_KEY_USER)
	If Len(aUserComponent(S_PASSWORD_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La contraseña del usuario está vacía."
		bIsCorrect = False
	End If
	If Len(aUserComponent(S_NAME_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del usuario está vacío."
		bIsCorrect = False
	End If
	If Len(aUserComponent(S_LAST_NAME_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El apellido del usuario está vacío."
		bIsCorrect = False
	End If
	If Len(aUserComponent(S_EMAIL_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El correo electrónico del usuario está vacío."
		bIsCorrect = False
	End If
	'If Not IsNumeric(aUserComponent(L_PERMISSIONS_USER)) Then aUserComponent(L_PERMISSIONS_USER) = 0
	'If Not IsNumeric(aUserComponent(L_PERMISSIONS2_USER)) Then aUserComponent(L_PERMISSIONS2_USER) = 0
	'If Not IsNumeric(aUserComponent(L_PERMISSIONS3_USER)) Then aUserComponent(L_PERMISSIONS3_USER) = 0
	'If Not IsNumeric(aUserComponent(L_PERMISSIONS4_USER)) Then aUserComponent(L_PERMISSIONS4_USER) = 0
	If Not IsNumeric(aUserComponent(L_PERMISSION_REPORTS_USER)) Then aUserComponent(L_PERMISSION_REPORTS_USER) = 0
	If Not IsNumeric(aUserComponent(L_PERMISSION_REPORTS2_USER)) Then aUserComponent(L_PERMISSION_REPORTS2_USER) = 0
	If Len(aUserComponent(S_PERMISSIONS_AREAS_USER)) = 0 Then aUserComponent(S_PERMISSIONS_AREAS_USER) = "-2"
	If Not IsNumeric(aUserComponent(L_PERMISSIONS_ZONE_USER)) Then aUserComponent(L_PERMISSIONS_ZONE_USER) = -1
	If Len(aUserComponent(S_BOSS_EMAIL_USER)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El correo electrónico del jefe inmediato del usuario está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aUserComponent(N_PROFILE_ID_USER)) Then aUserComponent(N_PROFILE_ID_USER) = -1
	If Not IsNumeric(aUserComponent(N_ACTIVE_USER)) Then aUserComponent(N_ACTIVE_USER) = 1
	If Not IsNumeric(aUserComponent(N_BLOCKED_USER)) Then aUserComponent(N_BLOCKED_USER) = 0
	If Not IsNumeric(aUserComponent(N_TECH_SUPPORT_USER)) Then aUserComponent(N_TECH_SUPPORT_USER) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del usuario contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "UserComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckUserInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayUserForm(oRequest, oADODBConnection, sAction, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an user from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aUserComponent
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayUserForm"
	Dim asAreas
	Dim iIndex
	Dim sTemp
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber
	Dim bEmptyArea

	If aUserComponent(N_ID_USER) <> -2 Then
		lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var asProfiles = new Array("
				Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "UserProfiles", "ProfileID", "ProfilePermissions, ProfilePermissions2, ProfilePermissions3, ProfilePermissions4, PermissionReports, PermissionReports2", "(ProfileID > -1)", "ProfileID", sErrorDescription)
			Response.Write "['-1', '0', '0', '0', '0', '0', '0']);" & vbNewLine

			Response.Write "function SetPermissionsForUser(sProfileID) {" & vbNewLine
				Response.Write "var oForm = document.UserFrm;" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "for (var i=0; i<asProfiles.length; i++)" & vbNewLine
						Response.Write "if (asProfiles[i][0] == sProfileID) {" & vbNewLine
							Response.Write "SendURLValuesToForm('UserPermissions=' + asProfiles[i][1] + '&UserPermissions2=' + asProfiles[i][2] + '&UserPermissions3=' + asProfiles[i][3] + '&UserPermissions4=' + asProfiles[i][4] + '&PermissionReports=' + asProfiles[i][5], oForm);" & vbNewLine
							Response.Write "SendURLValuesToForm('UserPermissions2b=' + asProfiles[i][2] + '&UserPermissions3b=' + asProfiles[i][3], oForm);" & vbNewLine
							Response.Write "SendURLValuesToForm('UserPermissions2g=' + asProfiles[i][2] + '&UserPermissions3g=' + asProfiles[i][3], oForm);" & vbNewLine
							Response.Write "SendURLValuesToForm('UserPermissionsProfile' + sProfileID + '=' + asProfiles[i][1], oForm);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of SetPermissionsForUser" & vbNewLine
			
			Response.Write "function ShowPermissionsForProfile(sProfileID) {" & vbNewLine
				Response.Write "var iMaxProfileID = 8;" & vbNewLine
				If B_ISSSTE Then
					Response.Write "for (var i=1; i<=iMaxProfileID; i++)" & vbNewLine
						Response.Write "HideDisplay(document.all['Section0' + i + 'Div']);" & vbNewLine
					Response.Write "if (parseInt(sProfileID) > 0)" & vbNewLine
						Response.Write "ShowDisplay(document.all['Section0' + sProfileID + 'Div']);" & vbNewLine
				End If
			Response.Write "} // End of ShowPermissionsForProfile" & vbNewLine

			Response.Write "function CheckUserFields(oForm) {" & vbNewLine
				Response.Write "var oField = null;" & vbNewLine
				Response.Write "var lUserPermissions4 = 0;" & vbNewLine

				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.UserName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre del usuario.');" & vbNewLine
							Response.Write "oForm.UserName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.UserLastName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el apellido del usuario.');" & vbNewLine
							Response.Write "oForm.UserLastName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.UserAccessKey.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la clave de acceso del usuario.');" & vbNewLine
							Response.Write "oForm.UserAccessKey.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						If B_PORTAL Then
							Response.Write "oForm.UserPassword.value = oForm.UserAccessKey.value;" & vbNewLine
							Response.Write "oForm.UserPwdConfirmation.value = oForm.UserAccessKey.value;" & vbNewLine
						Else
							Response.Write "if (oForm.UserPassword.value.length == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir la contraseña del usuario.');" & vbNewLine
								Response.Write "oForm.UserPassword.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (oForm.UserPassword.value != oForm.UserPwdConfirmation.value) {" & vbNewLine
								Response.Write "alert('La contraseña del usuario no coincide con la confirmación. Favor de introducirlas de nuevo.');" & vbNewLine
								Response.Write "oForm.UserPassword.value = '';" & vbNewLine
								Response.Write "oForm.UserPwdConfirmation.value = '';" & vbNewLine
								Response.Write "oForm.UserPassword.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
						Response.Write "if (oForm.UserEmail.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la cuenta de correo electrónico del usuario.');" & vbNewLine
							Response.Write "oForm.UserEmail.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.BossEmail.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la cuenta de correo electrónico del jefe inmediato del usuario.');" & vbNewLine
							Response.Write "oForm.BossEmail.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (parseInt(oForm.ProfileID.value) > 0) {" & vbNewLine
							Response.Write "oField = eval('oForm.UserPermissionsProfile' + oForm.ProfileID.value);" & vbNewLine
							Response.Write "if (oField)" & vbNewLine
								Response.Write "for (i=0; i<oField.length; i++)" & vbNewLine
									Response.Write "if (oField[i].checked)" & vbNewLine
										Response.Write "lUserPermissions4 += parseInt(oField[i].value);" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "lUserPermissions4 = -1;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "oForm.UserPermissions4.value = lUserPermissions4;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckUserFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""UserFrm"" ID=""UserFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckUserFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Users"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aUserComponent(N_ID_USER) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""UserName"" ID=""UserNameTxt"" SIZE=""26"" MAXLENGTH=""100"" VALUE=""" & aUserComponent(S_NAME_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido(s):&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""UserLastName"" ID=""UserLastNameTxt"" SIZE=""26"" MAXLENGTH=""100"" VALUE=""" & aUserComponent(S_LAST_NAME_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR"
					If aUserComponent(N_ID_USER) <> -2 Then Response.Write " STYLE=""display: none"""
				Response.Write ">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave de acceso:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""UserAccessKey"" ID=""UserAccessKeyTxt"" SIZE=""26"" MAXLENGTH=""120"" VALUE=""" & aUserComponent(S_ACCESS_KEY_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If B_PORTAL Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPassword"" ID=""UserPasswordHdn"" VALUE=""" & aUserComponent(S_ACCESS_KEY_USER) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPwdConfirmation"" ID=""UserPwdConfirmationHdn"" VALUE=""" & aUserComponent(S_ACCESS_KEY_USER) & """ />"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Contraseña:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""PASSWORD"" NAME=""UserPassword"" ID=""UserPasswordTxt"" SIZE=""26"" MAXLENGTH=""120"" VALUE=""" & aUserComponent(S_PASSWORD_USER) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Confirmación:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""PASSWORD"" NAME=""UserPwdConfirmation"" ID=""UserPwdConfirmationTxt"" SIZE=""26"" MAXLENGTH=""120"" VALUE=""" & aUserComponent(S_PASSWORD_USER) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">E-mail:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""UserEmail"" ID=""UserEmailTxt"" SIZE=""26"" MAXLENGTH=""100"" VALUE=""" & aUserComponent(S_EMAIL_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">E-mail jefe directo:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BossEmail"" ID=""BossEmailTxt"" SIZE=""26"" MAXLENGTH=""100"" VALUE=""" & aUserComponent(S_BOSS_EMAIL_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AdditionalEmail"" ID=""AdditionalEmailTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & aUserComponent(S_ADDITIONAL_EMAIL_USER) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
				If False Then
					Response.Write "¿Enviar copia de los mensajes del sistema a la cuenta de correo adicional?<BR />"
					Response.Write "<INPUT TYPE=""Radio"" NAME=""UserActive"" ID=""UserActiveRd"" VALUE=""1"""
						If aUserComponent(N_ACTIVE_USER) = 1 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " />No&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""UserActive"" ID=""UserActiveRd"" VALUE=""0"""
						If aUserComponent(N_ACTIVE_USER) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " />Sí<BR />"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserActive"" ID=""UserActiveHdn"" VALUE=""1"" />"
				End If

				Response.Write "¿El usuario terminó el curso de capacitación del sistema?<BR />"
				Response.Write "<INPUT TYPE=""Radio"" NAME=""UserBlocked"" ID=""UserBlockedRd"" VALUE=""0"""
					If aUserComponent(N_BLOCKED_USER) = 0 Then
						Response.Write " CHECKED=""1"""
					End If
				Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""UserBlocked"" ID=""UserBlockedRd"" VALUE=""1"""
					If aUserComponent(N_BLOCKED_USER) = 1 Then
						Response.Write " CHECKED=""1"""
					End If
				Response.Write " />No<BR /><BR />"

				Response.Write "¿El usuario tendrá acceso al sistema de soporte técnico?<BR />"
				Response.Write "<INPUT TYPE=""Radio"" NAME=""TechSupport"" ID=""TechSupportRd"" VALUE=""1"""
					If aUserComponent(N_TECH_SUPPORT_USER) = 1 Then
						Response.Write " CHECKED=""1"""
					End If
				Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""TechSupport"" ID=""TechSupportRd"" VALUE=""0"""
					If aUserComponent(N_TECH_SUPPORT_USER) = 0 Then
						Response.Write " CHECKED=""1"""
					End If
				Response.Write " />No<BR /><BR />"

				sTemp = "," & aUserComponent(S_PERMISSIONS_AREAS_USER) & ","
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Centro de trabajo del usuario:</B></FONT>"
				If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
					Response.Write "<BR /><INPUT TYPE=""CHECKBOX"" NAME=""AreaID1"" ID=""AreaID1"" VALUE=""-1"""
						If StrComp(aUserComponent(S_PERMISSIONS_AREAS_USER), "-1", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""if (this.checked) {HideDisplay(document.all['AreasDiv']); UncheckAllItemsFromCheckboxes(AreaID); } else {ShowDisplay(document.all['AreasDiv']);}"" /><B>TODOS</B>"
				Else
					sCondition = " And (AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
				End If
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""770"" HEIGHT=""1"" /><BR />"
				Response.Write "<DIV NAME=""AreasDiv"" ID=""AreasDiv"" STYLE=""width: 770px; height: 400px; overflow: auto;"
					If StrComp(aUserComponent(S_PERMISSIONS_AREAS_USER), "-1", vbBinaryCompare) = 0 Then Response.Write " display: none;"
				Response.Write """>"
					sErrorDescription = "No se pudieron obtener los registros de la base de datos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaCode, AreaName From Areas Where (AreaID>-1) And (ParentID=-1) And (EndDate=30000000) And (Active=1) " & sCondition & " Order By AreaCode, AreaName", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						asAreas = ""
						Do While Not oRecordset.EOF
							asAreas = asAreas & CStr(oRecordset.Fields("AreaID").Value) & SECOND_LIST_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & LIST_SEPARATOR
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						If Len(asAreas) > 0 Then asAreas = Left(asAreas, (Len(asAreas) - Len(LIST_SEPARATOR)))
						asAreas = Split(asAreas, LIST_SEPARATOR)
						For iIndex = 0 To UBound(asAreas)
							asAreas(iIndex) = Split(asAreas(iIndex), SECOND_LIST_SEPARATOR)
							Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""AreaID"" ID=""AreaID"" VALUE=""" & asAreas(iIndex)(0) & """"
								If InStr(1, sTemp, ("," & asAreas(iIndex)(0) & ","), vbBinaryCompare) > 0 Then Response.Write " CHECKED=""1"""
							Response.Write " onClick=""CheckAllChildNodesFrom" & asAreas(iIndex)(0) & "(AreaID, this.checked);"" "
							Response.Write " />" 
							Response.Write "<B>" & asAreas(iIndex)(1) & "</B><BR />"
							sTemp = Replace(sTemp, ("," & asAreas(iIndex)(0) & ","), ",")
							sErrorDescription = "No se pudieron obtener los registros de la base de datos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaCode, AreaName From Areas Where (AreaID>-1) And (ParentID=" & asAreas(iIndex)(0) & ") And (EndDate=30000000) And (Active=1) " & sCondition & " Order By AreaCode, AreaName", "UserComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							bEmptyArea = True
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									bEmptyArea = False
									Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""AreaID"" ID=""AreaID"" VALUE=""" & CStr(oRecordset.Fields("AreaID").Value) & """"
										If InStr(1, sTemp, ("," & CStr(oRecordset.Fields("AreaID").Value) & ","), vbBinaryCompare) > 0 Then Response.Write " CHECKED=""1"""
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & asAreas(iIndex)(0) & ", AreaID);}"" "
									Response.Write " />" 
									Response.Write CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & "<BR />"
									sTemp = Replace(sTemp, ("," & CStr(oRecordset.Fields("AreaID").Value) & ","), ",")
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								If Not bEmptyArea Then
									oRecordset.MoveFirst
									Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
										Response.Write "function CheckAllChildNodesFrom" & asAreas(iIndex)(0) & "(oCheckboxes, bCheck) {" & vbNewLine
											Response.Write "for (var i=0; i<oCheckboxes.length; i++) {" & vbNewLine
												Do While Not oRecordset.EOF
													Response.Write "if (oCheckboxes[i].value == " & CStr(oRecordset.Fields("AreaID").Value) & ")" & vbNewLine
														Response.Write "oCheckboxes[i].checked = bCheck;" & vbNewLine
													oRecordset.MoveNext
													If Err.number <> 0 Then Exit Do
												Loop
											Response.Write "}" & vbNewLine
										Response.Write "} // End of CheckAllChildNodesFrom " & asAreas(iIndex)(0) & vbNewLine
									Response.Write "//--></SCRIPT>" & vbNewLine
								End If
							End If
							If Err.number <> 0 Then Exit For
						Next
					End If
					'If InStr(1, sTemp, ",", vbBinaryCompare) = 1 Then sTemp = Replace(sTemp, ",", "", 1, 1, vbBinaryCompare)
					'If StrComp(Right(sTemp, Len(",")), ",", vbBinaryCompare) = 0  Then sTemp = Left(sTemp, (Len(sTemp) - Len(",")))
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & sTemp & """ />"
				Response.Write "</DIV>"
				Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""770"" HEIGHT=""1"" /><BR /><BR />"

			Response.Write "</FONT>"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Perfil: </B></FONT>"
			Response.Write "<SELECT NAME=""ProfileID"" ID=""ProfileIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SetPermissionsForUser(this.value); ShowPermissionsForProfile(this.value);"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "UserProfiles", "ProfileID", "ProfileName", "(ProfileID>=0)", "ProfileName", aUserComponent(N_PROFILE_ID_USER), "Ninguno;;;-1", sErrorDescription)
			Response.Write "</SELECT><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Permisos en el sistema:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""770"" HEIGHT=""1"" /><BR />"

			Response.Write "<DIV STYLE=""width: 770px; height: 205px; overflow: auto;"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UpdatePermissions"" ID=""UpdatePermissionsHdn"" VALUE="""" />"
			
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then
						Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions"" ID=""UserPermissionsChPm"" VALUE=""" & N_ADD_PERMISSIONS & """"
							If aUserComponent(L_PERMISSIONS_USER) And N_ADD_PERMISSIONS Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " /> Agregar registros<BR />"
					End If
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then
						Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions"" ID=""UserPermissionsChPm"" VALUE=""" & N_MODIFY_PERMISSIONS & """"
							If aUserComponent(L_PERMISSIONS_USER) And N_MODIFY_PERMISSIONS Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " /> Modificar registros<BR />"
					End If
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then
						Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions"" ID=""UserPermissionsChPm"" VALUE=""" & N_REMOVE_PERMISSIONS & """"
							If aUserComponent(L_PERMISSIONS_USER) And N_REMOVE_PERMISSIONS Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " /> Eliminar registros<BR />"
					End If

					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions"" ID=""UserPermissionsHdn"" VALUE=""" & (N_BUDGET_PERMISSIONS + N_AREAS_PERMISSIONS + N_POSITIONS_PERMISSIONS + N_JOBS_PERMISSIONS + N_EMPLOYEES_PERMISSIONS + N_EMPLOYEE_PAYROLL_PERMISSIONS + N_SADE_PERMISSIONS + N_PAYROLL_PERMISSIONS + N_PAYMENTS_PERMISSIONS + N_REPORTS_PERMISSIONS + N_CATALOGS_PERMISSIONS + N_TACO_PERMISSIONS) & """ />"
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions"" ID=""UserPermissionsHdn"" VALUE=""" & "," & N_BUDGET_PERMISSIONS & "," & N_AREAS_PERMISSIONS & "," & N_POSITIONS_PERMISSIONS & "," & N_JOBS_PERMISSIONS & "," & N_EMPLOYEES_PERMISSIONS & "," & N_EMPLOYEE_PAYROLL_PERMISSIONS & "," & N_SADE_PERMISSIONS & "," & N_PAYROLL_PERMISSIONS & "," & N_PAYMENTS_PERMISSIONS & "," & N_REPORTS_PERMISSIONS & "," & N_CATALOGS_PERMISSIONS & "," & N_TACO_PERMISSIONS & """ />"
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions2"" ID=""UserPermissions2Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS2_USER) & """ />"
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions3"" ID=""UserPermissions3Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS3_USER) & """ />"
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions4"" ID=""UserPermissions4Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS4_USER) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports"" ID=""PermissionReportsHdn"" VALUE=""" & aUserComponent(L_PERMISSION_REPORTS_USER) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports2"" ID=""PermissionReports2Hdn"" VALUE=""" & aUserComponent(L_PERMISSION_REPORTS2_USER) & """ />"

					Response.Write "<DIV NAME=""Section01Div"" ID=""Section01Div"" STYLE=""display: none"">"
						If aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_01_PERMISSIONS4 And False Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_PERMISSIONS4 & """"
								If aUserComponent(L_PERMISSIONS4_USER) And N_01_PERMISSIONS4 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_15_PERMISSIONS4 & ", UserPermissionsProfile1);} else {SetCheckboxesValue(" & N_15_PERMISSIONS4 & ", UserPermissionsProfile1);}"" /> Administración de plazas<BR />"
						End If
							If False And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_15_PERMISSIONS4 & """"
									If aUserComponent(L_PERMISSIONS4_USER) And N_15_PERMISSIONS4 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_PERMISSIONS4 & ", UserPermissionsProfile1);}"" /> Modificación de plazas<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AsignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AsignacionDeNumeroDeEmpleado & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AsignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Asignación de número de empleado<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ConsultaDePersonal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_ConsultaDePersonal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_ConsultaDePersonal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Consulta de personal<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AdministracionDePersonal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AdministracionDePersonal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AdministracionDePersonal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_01_ModificacionDePersonal & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_ModificacionDeAntiguedades & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_ValidacionDeMovimientos & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_AplicacionDeMovimientos & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_101NuevoIngreso & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_AltaDeHonorarios & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_103_AltaPorInterinato & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_106_AltaPorReingreso & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_106_AltaParaOcuparPuestoDeConfianza & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_107_AltaPorReinstalacion & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_130_AltaPorReanudacionDeLabores & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_109_RetornoALaVidaLaboral & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_210_CambioDePlazaMismaAdscripcion & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_212_CambioDeAdscripcionConPlaza & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_213_CambioDeAdscripcionSinPlaza & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_220_PermutaDePlazas & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_CambioDeDatosDelEmpleado & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_000_ReasignacionDeNumeroDeEmpleado & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_220_InclusionDeRiesgosProfesionales & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_220_TurnoOpcional_Concepto_07 & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_220_PercepcionAdicional_Concepto_08 & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_CambioDeHonorarios & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaDeRegistrosVigentes & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_441_LicenciaCGSPorComisionSindical & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_446_LicenciaCGSPorTramiteDePension & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_448_LicenciaCGSPorContraerMatrimonio & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_592_LicenciaSGSPorAsuntosParticulares & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_593_LicenciaSGSPorComisionSindical & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_310_BajaDePersonalDeHonorarios & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_340_BajaPorRenuncia & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_341_BajaPorDefuncion & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_342_BajaPorCese & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_343_BajaPorIncapacidadTotalYPermanente & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_344_BajaPorPension & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_345_BajaPorJubilacion & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_346_BajaPorInterinato & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_CambioDeDatosDelEmpleado & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_348_BajaPorTerminoAlPuestoDeConfianza & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_349_BajaPorSancionAdministrativa & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_350_BajaPorSancion & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaPorTerminoDeNombramiento & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaPorTerminoDeProvisionalidad & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_TerminoDeConvenio & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaPorNoTomarPosesionDelPuesto & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaPorIncurrirEnFaltaAdministrativa & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_370_BajaPorRenunciaEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_371_BajaPorDefuncionEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_372_BajaPorCeseEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_374_BajaPorPensionEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_375_BajaPorJubilacionEnLSGS & ", UserPermissionsProfile1); } else { SetCheckboxesValue(" & N_01_ModificacionDePersonal & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_ModificacionDeAntiguedades & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_ValidacionDeMovimientos & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_AplicacionDeMovimientos & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_101NuevoIngreso & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_AltaDeHonorarios & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_103_AltaPorInterinato & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_106_AltaPorReingreso & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_106_AltaParaOcuparPuestoDeConfianza & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_107_AltaPorReinstalacion & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_130_AltaPorReanudacionDeLabores & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_109_RetornoALaVidaLaboral & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_210_CambioDePlazaMismaAdscripcion & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_212_CambioDeAdscripcionConPlaza & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_213_CambioDeAdscripcionSinPlaza & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_220_PermutaDePlazas & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_CambioDeDatosDelEmpleado & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_000_ReasignacionDeNumeroDeEmpleado & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_220_InclusionDeRiesgosProfesionales & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_220_TurnoOpcional_Concepto_07 & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_220_PercepcionAdicional_Concepto_08 & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_CambioDeHonorarios & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaDeRegistrosVigentes & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_441_LicenciaCGSPorComisionSindical & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_446_LicenciaCGSPorTramiteDePension & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_448_LicenciaCGSPorContraerMatrimonio & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_592_LicenciaSGSPorAsuntosParticulares & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_593_LicenciaSGSPorComisionSindical & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_310_BajaDePersonalDeHonorarios & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_340_BajaPorRenuncia & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_341_BajaPorDefuncion & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_342_BajaPorCese & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_343_BajaPorIncapacidadTotalYPermanente & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_344_BajaPorPension & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_345_BajaPorJubilacion & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_346_BajaPorInterinato & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_CambioDeDatosDelEmpleado & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_348_BajaPorTerminoAlPuestoDeConfianza & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_349_BajaPorSancionAdministrativa & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_350_BajaPorSancion & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaPorTerminoDeNombramiento & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaPorTerminoDeProvisionalidad & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_TerminoDeConvenio & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaPorNoTomarPosesionDelPuesto & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaPorIncurrirEnFaltaAdministrativa & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_370_BajaPorRenunciaEnLSGS & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_371_BajaPorDefuncionEnLSGS & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_372_BajaPorCeseEnLSGS & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_374_BajaPorPensionEnLSGS & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_375_BajaPorJubilacionEnLSGS & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_351_BajaPorTransicionDLaTerminacionDlaRelacionLaboralConInstituto & ", UserPermissionsProfile1); }"" /> Administración de personal<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDePersonal & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_ModificacionDePersonal & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_ModificacionDePersonal & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Modificación de personal<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDeAntiguedades & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_ModificacionDeAntiguedades & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_ModificacionDeAntiguedades & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Modificación de antigüedades<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_ValidacionDeMovimientos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Validación de movimientos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AplicacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AplicacionDeMovimientos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AplicacionDeMovimientos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Aplicación de movimientos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_101NuevoIngreso & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_101NuevoIngreso & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_101NuevoIngreso & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 101 Nuevo ingreso<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AltaDeHonorarios & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AltaDeHonorarios & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AltaDeHonorarios & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Alta de honorarios<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_103_AltaPorInterinato & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_103_AltaPorInterinato & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_103_AltaPorInterinato & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 103 Alta por interinato<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_106_AltaPorReingreso & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_106_AltaPorReingreso & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_106_AltaPorReingreso & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 106 Alta por reingreso<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_106_AltaParaOcuparPuestoDeConfianza & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_106_AltaParaOcuparPuestoDeConfianza & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_106_AltaParaOcuparPuestoDeConfianza & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 106 Alta para ocupar puesto de confianza<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_107_AltaPorReinstalacion & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_107_AltaPorReinstalacion & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_107_AltaPorReinstalacion & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 107 Alta por reinstalación<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_130_AltaPorReanudacionDeLabores & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_130_AltaPorReanudacionDeLabores & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_130_AltaPorReanudacionDeLabores & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 130 Alta por reanudación de labores<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_109_RetornoALaVidaLaboral & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_109_RetornoALaVidaLaboral & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_109_RetornoALaVidaLaboral & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 109 Retorno a la vida laboral<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_380_BajaPorTransicionDLaTerminacionDlaRelacionLaboralConInstituto & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_380_BajaPorTransicionDLaTerminacionDlaRelacionLaboralConInstituto & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_380_BajaPorTransicionDLaTerminacionDlaRelacionLaboralConInstituto & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 109 Retorno a la vida laboral<BR />"
							End If							
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_210_CambioDePlazaMismaAdscripcion & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_210_CambioDePlazaMismaAdscripcion & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_210_CambioDePlazaMismaAdscripcion & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 210 Cambio de plaza misma adscripción<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_212_CambioDeAdscripcionConPlaza & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_212_CambioDeAdscripcionConPlaza & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_212_CambioDeAdscripcionConPlaza & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 212 Cambio de adscripción con plaza / 219 Cambio de servicio / 219 Cambio de turno<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_213_CambioDeAdscripcionSinPlaza & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_213_CambioDeAdscripcionSinPlaza & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_213_CambioDeAdscripcionSinPlaza & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 213 Cambio de adscripción sin plaza<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_PermutaDePlazas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_220_PermutaDePlazas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_220_PermutaDePlazas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 218. Permuta de plazas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_CambioDeDatosDelEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_CambioDeDatosDelEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_CambioDeDatosDelEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 220. Cambio de datos del empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_000_ReasignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1hPm"" VALUE=""" & N_01_000_ReasignacionDeNumeroDeEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_000_ReasignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 000 Reasignación de número de empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_InclusionDeRiesgosProfesionales & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_220_InclusionDeRiesgosProfesionales & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_220_InclusionDeRiesgosProfesionales & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 222 Inclusión de Riesgos profesionales<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_TurnoOpcional_Concepto_07 & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_220_TurnoOpcional_Concepto_07 & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_220_TurnoOpcional_Concepto_07 & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 221. Turno opcional (Concepto 07)<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_PercepcionAdicional_Concepto_08 & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_220_PercepcionAdicional_Concepto_08 & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_220_PercepcionAdicional_Concepto_08 & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 221 Percepción adicional (Concepto 08)<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_CambioDeHonorarios & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_CambioDeHonorarios & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_CambioDeHonorarios & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Cambio de Honorarios (Concepto 11)<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaDeRegistrosVigentes & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaDeRegistrosVigentes & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaDeRegistrosVigentes & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Baja de registros vigentes<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_441_LicenciaCGSPorComisionSindical & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_441_LicenciaCGSPorComisionSindical & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_441_LicenciaCGSPorComisionSindical & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 441 Licencia con goce de sueldo por Comisión sindical<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_446_LicenciaCGSPorTramiteDePension & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_446_LicenciaCGSPorTramiteDePension & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_446_LicenciaCGSPorTramiteDePension & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 446 Licencia con goce de sueldo por trámite de pensión<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_448_LicenciaCGSPorContraerMatrimonio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_448_LicenciaCGSPorContraerMatrimonio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_448_LicenciaCGSPorContraerMatrimonio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 448 Licencia con goce de sueldo por contraer matrimonio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 449 Licencia con goce de sueldo por fallecimiento de familiar en primer grado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 451 Licencia con goce de sueldo por otorgamiento de beca<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 452 Licencia con goce de sueldo por práctica de servicio social<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_592_LicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_592_LicenciaSGSPorAsuntosParticulares & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_592_LicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 592 Licencia sin goce de sueldo por asuntos particulares<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_593_LicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_593_LicenciaSGSPorComisionSindical & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_593_LicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 593 Licencia sin goce de sueldo por comisión sindical<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 594 Licencia sin goce de sueldo por otorgamiento de beca<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 595 Licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del instituto<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 596 Licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 597 Licencia sin goce de sueldo por práctica de servicio social<BR />"
							End If
							If (False And ((StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_468_ProrrogaDeLicenciaConSueldoPorComisionsindical & ",", vbBinaryCompare) > 0))) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_468_ProrrogaDeLicenciaConSueldoPorComisionsindical & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_468_ProrrogaDeLicenciaConSueldoPorComisionsindical & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 469 Prórroga de licencia con goce de sueldo por otorgamiento de beca<BR />"
							End If							
							If (False And ((StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0))) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 469 Prórroga de licencia con goce de sueldo por otorgamiento de beca<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 570 Prórroga de licencia sin goce de sueldo por comisión sindical<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 571 Prórroga de licencia sin goce de sueldo por otorgamiento de beca<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 572 Prórroga de licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del Instituto<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 573 Prórroga de licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 574 Prórroga de licencia sin goce de sueldo por asuntos particulares<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_310_BajaDePersonalDeHonorarios & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_310_BajaDePersonalDeHonorarios & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_310_BajaDePersonalDeHonorarios & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 310 Baja de personal de honorarios<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_340_BajaPorRenuncia & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_340_BajaPorRenuncia & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_340_BajaPorRenuncia & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 340 Baja por renuncia<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_341_BajaPorDefuncion & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_341_BajaPorDefuncion & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_341_BajaPorDefuncion & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 341 Baja por defunción<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_342_BajaPorCese & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_342_BajaPorCese & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_342_BajaPorCese & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 342 Baja por cese<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_343_BajaPorIncapacidadTotalYPermanente & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_343_BajaPorIncapacidadTotalYPermanente & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_343_BajaPorIncapacidadTotalYPermanente & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 343 Baja por incapacidad total y permanente<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_344_BajaPorPension & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_344_BajaPorPension & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_344_BajaPorPension & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 344 Baja por pensión<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_345_BajaPorJubilacion & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_345_BajaPorJubilacion & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_345_BajaPorJubilacion & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 345 Baja por jubilación<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_346_BajaPorInterinato & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_346_BajaPorInterinato & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_346_BajaPorInterinato & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 350 Baja por interinato<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_348_BajaPorTerminoAlPuestoDeConfianza & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_348_BajaPorTerminoAlPuestoDeConfianza & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_348_BajaPorTerminoAlPuestoDeConfianza & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 348 Baja por término al puesto de confianza<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_349_BajaPorSancionAdministrativa & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_349_BajaPorSancionAdministrativa & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_349_BajaPorSancionAdministrativa & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 349 Baja por sanción administrativa<BR />"
							End If
							If (False And ((StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_350_BajaPorSancion & ",", vbBinaryCompare) > 0))) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_350_BajaPorSancion & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_350_BajaPorSancion & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 350 Baja por sanción<BR />"
							End If
							If (False And ((StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorTerminoDeNombramiento & ",", vbBinaryCompare) > 0))) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaPorTerminoDeNombramiento & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaPorTerminoDeNombramiento & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Baja por término de nombramiento<BR />"
							End If
							If (False And ((StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorTerminoDeProvisionalidad & ",", vbBinaryCompare) > 0))) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaPorTerminoDeProvisionalidad & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaPorTerminoDeProvisionalidad & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> Baja por término de provisionalidad<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_TerminoDeConvenio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_TerminoDeConvenio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_TerminoDeConvenio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 346 Término de convenio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorNoTomarPosesionDelPuesto & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaPorNoTomarPosesionDelPuesto & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaPorNoTomarPosesionDelPuesto & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 347 Baja por no tomar posesión del puesto<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaPorIncurrirEnFaltaAdministrativa & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 348 Baja por incurrir en falta administrativa antes de adquirir la inamovilidad<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_370_BajaPorRenunciaEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_370_BajaPorRenunciaEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_370_BajaPorRenunciaEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 370 Baja por renuncia con estatus de licencia sin goce de sueldo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_371_BajaPorDefuncionEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_371_BajaPorDefuncionEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_371_BajaPorDefuncionEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 371 Baja por defunción con estatus de licencia sin goce de sueldo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_372_BajaPorCeseEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_372_BajaPorCeseEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_372_BajaPorCeseEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 372 Baja por cese con estatus de licencia sin goce de sueldo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 373 Baja por incapacidad total y permanente con estatus de licencia sin goce de sueldo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_374_BajaPorPensionEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_374_BajaPorPensionEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_374_BajaPorPensionEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 374 Baja por pensión con estatus de licencia sin goce de sueldo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_375_BajaPorJubilacionEnLSGS & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_375_BajaPorJubilacionEnLSGS & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_375_BajaPorJubilacionEnLSGS & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_AdministracionDePersonal & ", UserPermissionsProfile1);}"" /> 375 Baja por jubilación con estatus de licencia sin goce de sueldo<BR />"
							End If

						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Aguinaldos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_Aguinaldos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_Aguinaldos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Aguinaldos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AcumuladosAnuales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AcumuladosAnuales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AcumuladosAnuales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Acumulados anuales<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_ReclamosDePagoPorAjustesYDeducciones & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_01_C9_DevolucionesNoGravables & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_72_OtrasDeducciones & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_RegistroDeReclamos & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_RevisionDeNominas & ", UserPermissionsProfile1); UncheckCheckboxesValue(" & N_01_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile1); } else {SetCheckboxesValue(" & N_01_C9_DevolucionesNoGravables & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_72_OtrasDeducciones & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_RegistroDeReclamos & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_RevisionDeNominas & ", UserPermissionsProfile1); SetCheckboxesValue(" & N_01_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile1);}"" /> Reclamos de pago por ajustes y deducciones<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_C9_DevolucionesNoGravables & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> C9. Devoluciones no gravables<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_71_DeduccionPorCobroDeSueldosIndebidos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> 71. Deducción por cobro de sueldos indebidos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_72_OtrasDeducciones & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_72_OtrasDeducciones & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_72_OtrasDeducciones & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> 72. Otras deducciones<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_RegistroDeReclamos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> Registro de reclamos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RevisionDeNominas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_RevisionDeNominas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_RevisionDeNominas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> Revisión de nóminas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_BajaDeRegistrosVigentesReclamos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_01_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile1);}"" /> Baja de registros vigentes<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Catalogos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_Catalogos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_Catalogos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Catálogos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_TableroDeControl & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_TableroDeControl & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_TableroDeControl & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Tablero de control de procesos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Ventanilla única<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_01_AcreedoresDeLosEmpleados & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Acreedores de los empleados<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile1"" ID=""UserPermissionsProfile1ChPm"" VALUE=""" & N_31_PERMISSIONS4 & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_01_RevisionDeNominas & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
					Response.Write "</DIV>"
					Response.Write "<DIV NAME=""Section02Div"" ID=""Section02Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConsultaDePersonal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ConsultaDePersonal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ConsultaDePersonal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Consulta de personal<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_02_SI_SeguroDeSeparacionIndividualizado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_BajaDeRegistrosVigentesSIyAE & ", UserPermissionsProfile2);} else {SetCheckboxesValue(" & N_02_SI_SeguroDeSeparacionIndividualizado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_BajaDeRegistrosVigentesSIyAE & ", UserPermissionsProfile2);}"" /> SI. Seguro de separación y AE. Seguro adicional de separación individualizado<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_SI_SeguroDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_SI_SeguroDeSeparacionIndividualizado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_SI_SeguroDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ", UserPermissionsProfile2);}"" /> SI. Seguro de separación individualizado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ", UserPermissionsProfile2);}"" /> AE. Seguro adicional de separación individualizado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_BajaDeRegistrosVigentesSIyAE & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_BajaDeRegistrosVigentesSIyAE & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_BajaDeRegistrosVigentesSIyAE & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ", UserPermissionsProfile2);}"" /> Baja de registros vigentes<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CertificacionesYArchivo & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_CertificacionesYArchivo & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_CertificacionesYArchivo & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_02_RegistroDeReclamosCyA & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ActualizacionDeAntiguedades & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_AntiguedadFederal & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_AntiguedadParaUnEmpleado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ReporteDeAntiguedades & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ValidacionDeNomina1oDeOctubre & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_HojaUnicaDeServicio & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_EntregasDeHojasUnicasDeServicio & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ConstanciaDeServicioActivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ConstanciaDeDescuento & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RevisionDeNominasCyA & ", UserPermissionsProfile2);} else {SetCheckboxesValue(" & N_02_RegistroDeReclamosCyA & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ActualizacionDeAntiguedades & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_AntiguedadFederal & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_AntiguedadParaUnEmpleado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ReporteDeAntiguedades & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ValidacionDeNomina1oDeOctubre & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_HojaUnicaDeServicio & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_EntregasDeHojasUnicasDeServicio & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ConstanciaDeServicioActivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ConstanciaDeDescuento & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RevisionDeNominasCyA & ", UserPermissionsProfile2);}"" /> Certificaciones y archivo<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDeReclamosCyA & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Registro de reclamos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> D2. Exceso de incapacidades y licencias médicas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RevisionDeNominasCyA & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissions3ChPm"" VALUE=""" & N_02_RevisionDeNominasCyA & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RevisionDeNominasCyA & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Revisión de nóminas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ActualizacionDeAntiguedades & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ActualizacionDeAntiguedades & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ActualizacionDeAntiguedades & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Actualización de antigüedades<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AntiguedadFederal & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AntiguedadFederal & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AntiguedadFederal & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Antigüedad federal<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AntiguedadParaUnEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AntiguedadParaUnEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AntiguedadParaUnEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Antigüedad para un empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeAntiguedades & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ReporteDeAntiguedades & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ReporteDeAntiguedades & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Reporte de antigüedades<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ValidacionDeNomina1oDeOctubre & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ValidacionDeNomina1oDeOctubre & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ValidacionDeNomina1oDeOctubre & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Validación de nómina 1o de Octubre<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_HojaUnicaDeServicio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_HojaUnicaDeServicio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_HojaUnicaDeServicio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Hoja única de servicio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_EntregasDeHojasUnicasDeServicio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Entregas de hojas únicas de servicio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConstanciaDeServicioActivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ConstanciaDeServicioActivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ConstanciaDeServicioActivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Constancia de servicio activo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConstanciaDeDescuento & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ConstanciaDeDescuento & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ConstanciaDeDescuento & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_CertificacionesYArchivo & ", UserPermissionsProfile2);}"" /> Constancia de descuento<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_TercerosInstitucionales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_TercerosInstitucionales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_TercerosInstitucionales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_02_CargaDeDiscosDeTerceros & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RegistroEnLineaDeTercerosInstitucionales & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_GeneracionDeArchivoRepcsi & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RegistroDeCalificacionDeEmpleados & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ReporteDeCalificacionDeEmpleados & ", UserPermissionsProfile2); } else { SetCheckboxesValue(" & N_02_CargaDeDiscosDeTerceros & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RegistroEnLineaDeTercerosInstitucionales & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_GeneracionDeArchivoRepcsi & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RegistroDeCalificacionDeEmpleados & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ReporteDeCalificacionDeEmpleados & ", UserPermissionsProfile2); }"" /> Terceros institucionales<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CargaDeDiscosDeTerceros & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_CargaDeDiscosDeTerceros & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_CargaDeDiscosDeTerceros & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Carga de discos de terceros<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Aplicación de registros cargados por cada archivo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroEnLineaDeTercerosInstitucionales & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroEnLineaDeTercerosInstitucionales & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroEnLineaDeTercerosInstitucionales & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Registro en línea de terceros institucionales<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ReporteDeCargaDesdeArchivosDeTerceros & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Reporte de carga desde archivos de terceros<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_GeneracionDeArchivoRepcsi & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_GeneracionDeArchivoRepcsi & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_GeneracionDeArchivoRepcsi & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Generación de archivo Repcsi<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDeCalificacionDeEmpleados & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Registro de calificación de empleados<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ReporteDeCalificacionDeEmpleados & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ReporteDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_TercerosInstitucionales & ", UserPermissionsProfile2);}"" /> Reporte de calificación de empleados<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_PrestacionesEIncidencias & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_02_RegistroDeReclamos & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RevisionDeNominas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RegistroDeConceptosDeEmpleado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_05_CompensacionesPorAntiguedad & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_09_RemuneracionPorHorasExtraordinarias & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_14_PrimasDominicales & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_19_BecasParaLosHijosDeLosTrabajadores & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_20_AyudaDeAnteojos & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_28_EstimuloALaProductividad & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_43_AyudaImpresionDeTesis & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_49_PremioTrabajadorDelMes & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_22_Premio10DeMayo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_67_ApoyoDeportivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_67_CuotaDeportivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_C3_PremiosEstimulosYRecompensas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_16_DevolucionesPorDeduccionesIndebidas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_71_DevolucionesNoExcentas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_72_OtrasDeducciones & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_61_IndicadorComisionDeAuxilio & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_65_BajaDeSeguroColectivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_77_FondoDeAhorroCapitalizableFonac & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_19_PERMISSIONS2 & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_76_AjusteFonac & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_7s_AhorroSolidario & ", UserPermissionsProfile2); } else { SetCheckboxesValue(" & N_02_RegistroDeReclamos & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RevisionDeNominas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RegistroDeConceptosDeEmpleado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_05_CompensacionesPorAntiguedad & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_09_RemuneracionPorHorasExtraordinarias & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_14_PrimasDominicales & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_19_BecasParaLosHijosDeLosTrabajadores & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_20_AyudaDeAnteojos & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_28_EstimuloALaProductividad & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_43_AyudaImpresionDeTesis & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_49_PremioTrabajadorDelMes & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_22_Premio10DeMayo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_67_ApoyoDeportivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_67_CuotaDeportivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_C3_PremiosEstimulosYRecompensas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_16_DevolucionesPorDeduccionesIndebidas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_71_DevolucionesNoExcentas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_72_OtrasDeducciones & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_61_IndicadorComisionDeAuxilio & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_65_BajaDeSeguroColectivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_77_FondoDeAhorroCapitalizableFonac & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_76_AjusteFonac & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_7s_AhorroSolidario & ", UserPermissionsProfile2); }"" /> Prestaciones e incidencias<BR />"
'							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_02_RegistroDeReclamos & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RevisionDeNominas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_RegistroDeConceptosDeEmpleado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_05_CompensacionesPorAntiguedad & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_09_RemuneracionPorHorasExtraordinarias & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_14_PrimasDominicales & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_19_BecasParaLosHijosDeLosTrabajadores & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_20_AyudaDeAnteojos & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_28_EstimuloALaProductividad & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_43_AyudaImpresionDeTesis & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_49_PremioTrabajadorDelMes & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_22_Premio10DeMayo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_41_Premio_antiguedad_25_y_30_anios & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_67_ApoyoDeportivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_67_CuotaDeportivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_C3_PremiosEstimulosYRecompensas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_16_DevolucionesPorDeduccionesIndebidas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_71_DevolucionesNoExcentas & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_72_OtrasDeducciones & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_61_IndicadorComisionDeAuxilio & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_65_BajaDeSeguroColectivo & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_77_FondoDeAhorroCapitalizableFonac & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_19_PERMISSIONS2 & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_76_AjusteFonac & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_7s_AhorroSolidario & ", UserPermissionsProfile2); } else { SetCheckboxesValue(" & N_02_RegistroDeReclamos & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RevisionDeNominas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_RegistroDeConceptosDeEmpleado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_05_CompensacionesPorAntiguedad & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_09_RemuneracionPorHorasExtraordinarias & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_14_PrimasDominicales & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_19_BecasParaLosHijosDeLosTrabajadores & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_20_AyudaDeAnteojos & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_28_EstimuloALaProductividad & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_43_AyudaImpresionDeTesis & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_49_PremioTrabajadorDelMes & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_22_Premio10DeMayo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_41_Premio_antiguedad_25_y_30_anios & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_67_ApoyoDeportivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_67_CuotaDeportivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_C3_PremiosEstimulosYRecompensas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_16_DevolucionesPorDeduccionesIndebidas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_71_DevolucionesNoExcentas & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_72_OtrasDeducciones & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_61_IndicadorComisionDeAuxilio & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_65_BajaDeSeguroColectivo & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_77_FondoDeAhorroCapitalizableFonac & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_76_AjusteFonac & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_7s_AhorroSolidario & ", UserPermissionsProfile2); }"" /> Prestaciones e incidencias<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDeReclamos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> Registro de reclamos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RevisionDeNominas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RevisionDeNominas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RevisionDeNominas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> Revisión de nóminas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeConceptosDeEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDeConceptosDeEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDeConceptosDeEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> Registro de conceptos de empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_05_CompensacionesPorAntiguedad & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_05_CompensacionesPorAntiguedad & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_05_CompensacionesPorAntiguedad & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 05. Compensaciones por antigüedad<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_09_RemuneracionPorHorasExtraordinarias & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_09_RemuneracionPorHorasExtraordinarias & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_09_RemuneracionPorHorasExtraordinarias & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 09. Remuneración por horas extraordinarias<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_14_PrimasDominicales & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_14_PrimasDominicales & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_14_PrimasDominicales & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 14. Primas dominicales<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_19_BecasParaLosHijosDeLosTrabajadores & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_19_BecasParaLosHijosDeLosTrabajadores & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_19_BecasParaLosHijosDeLosTrabajadores & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 19. Becas para los hijos de los trabajadores<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_20_AyudaDeAnteojos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_20_AyudaDeAnteojos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_20_AyudaDeAnteojos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 20. Ayuda de anteojos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_28_EstimuloALaProductividad & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_28_EstimuloALaProductividad & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_28_EstimuloALaProductividad & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 28. Estímulo a la productividad, calidad y eficacia para personal médico y de enfermería<BR />"
							End If
							
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_41_Premio_antiguedad_25_y_30_anios & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_41_Premio_antiguedad_25_y_30_anios & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_41_Premio_antiguedad_25_y_30_anios & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 41. Premio antigüedad 25 y 30 años<BR />"
							End If	
							
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 42. Ayuda por muerte de familiar en primer grado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_43_AyudaImpresionDeTesis & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_43_AyudaImpresionDeTesis & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_43_AyudaImpresionDeTesis & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 43. Ayuda impresión de tesis<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_49_PremioTrabajadorDelMes & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_49_PremioTrabajadorDelMes & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_49_PremioTrabajadorDelMes & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 49. Premio trabajador del mes<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_22_Premio10DeMayo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_22_Premio10DeMayo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_22_Premio10DeMayo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 22. Premio 10 de Mayo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_67_ApoyoDeportivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_67_ApoyoDeportivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_67_ApoyoDeportivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> AD. Apoyo al Deporte<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_67_CuotaDeportivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_67_CuotaDeportivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_67_CuotaDeportivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 67. Cuota deportivo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> C2. Jornada nocturna adicional por día festivo (acumulada)<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_C3_PremiosEstimulosYRecompensas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_C3_PremiosEstimulosYRecompensas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_C3_PremiosEstimulosYRecompensas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> C3. Premios, estímulos y recompensas (recompensa del sistema de evaluación del desempeño)<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_16_DevolucionesPorDeduccionesIndebidas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_16_DevolucionesPorDeduccionesIndebidas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_16_DevolucionesPorDeduccionesIndebidas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 16. Devoluciones por deducciones indebidas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_71_DevolucionesNoExcentas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_71_DevolucionesNoExcentas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_71_DevolucionesNoExcentas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 71. Devoluciones no excentas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_72_OtrasDeducciones & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_72_OtrasDeducciones & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_72_OtrasDeducciones & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 72. Otras deducciones<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_61_IndicadorComisionDeAuxilio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_61_IndicadorComisionDeAuxilio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_61_IndicadorComisionDeAuxilio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 61. Indicador comisión de auxilio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_65_BajaDeSeguroColectivo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_65_BajaDeSeguroColectivo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_65_BajaDeSeguroColectivo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 65. Baja de seguro colectivo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_77_FondoDeAhorroCapitalizableFonac & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_77_FondoDeAhorroCapitalizableFonac & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_77_FondoDeAhorroCapitalizableFonac & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 77. Fondo de ahorro capitalizable FONAC<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_76_AjusteFonac & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_76_AjusteFonac & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_76_AjusteFonac & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 76. Ajuste FONAC<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_7s_AhorroSolidario & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_7s_AhorroSolidario & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_7s_AhorroSolidario & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PrestacionesEIncidencias & ", UserPermissionsProfile2);}"" /> 7S. Ahorro solidario<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_BajaDePrestacionesVigentes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_BajaDePrestacionesVigentes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_BajaDePrestacionesVigentes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Baja de Prestaciones vigentes<BR />"
						End If
'							If aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_03_PERMISSIONS4 Then
'								Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_03_PERMISSIONS4 & """"
'									If aUserComponent(L_PERMISSIONS4_USER) And N_03_PERMISSIONS4 Then
'										Response.Write " CHECKED=""1"""
'									End If
'									Response.Write " /> Antigüedades<BR />"
'							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_PensionAlimenticia & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_PensionAlimenticia & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_PensionAlimenticia & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_02_RegistroDePensionAlimenticia & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_CatalogoDeTiposDePensionAlimenticia & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_AdeudoPensionAlimenticia & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_ReporteDeEmpleadosConPensionesAlimenticias & ", UserPermissionsProfile2);} else {SetCheckboxesValue(" & N_02_RegistroDePensionAlimenticia & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_CatalogoDeTiposDePensionAlimenticia & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_AdeudoPensionAlimenticia & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_ReporteDeEmpleadosConPensionesAlimenticias & ", UserPermissionsProfile2);}"" /> Pensión alimenticia<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDePensionAlimenticia & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDePensionAlimenticia & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDePensionAlimenticia & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PensionAlimenticia & ", UserPermissionsProfile2);}"" /> Pensión alimenticia<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CatalogoDeTiposDePensionAlimenticia & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_CatalogoDeTiposDePensionAlimenticia & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_CatalogoDeTiposDePensionAlimenticia & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PensionAlimenticia & ", UserPermissionsProfile2);}"" /> Catálogo de tipos de pensión alimenticia<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdeudoPensionAlimenticia & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AdeudoPensionAlimenticia & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AdeudoPensionAlimenticia & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PensionAlimenticia & ", UserPermissionsProfile2);}"" /> Adeudo pensión alimenticia<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PensionAlimenticia & ", UserPermissionsProfile2);}"" /> Reporte de beneficiarios de pensiones alimenticias por empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeEmpleadosConPensionesAlimenticias & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_ReporteDeEmpleadosConPensionesAlimenticias & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_ReporteDeEmpleadosConPensionesAlimenticias & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_PensionAlimenticia & ", UserPermissionsProfile2);}"" /> Reporte de empleados con pensiones alimenticias<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AcreedoresDeLosEmpleados & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Acreedores de los empleados<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_MatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_MatrizDeRiesgosProfesionales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_MatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_02_AdministracionDeMatrizDeRiesgos & ", UserPermissionsProfile2); UncheckCheckboxesValue(" & N_02_CargaDeMatrizDeRiesgosProfesionales & ", UserPermissionsProfile2);} else {SetCheckboxesValue(" & N_02_AdministracionDeMatrizDeRiesgos & ", UserPermissionsProfile2); SetCheckboxesValue(" & N_02_CargaDeMatrizDeRiesgosProfesionales & ", UserPermissionsProfile2);}"" /> Matriz de riesgos profesionales<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministracionDeMatrizDeRiesgos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2bChPm"" VALUE=""" & N_02_AdministracionDeMatrizDeRiesgos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AdministracionDeMatrizDeRiesgos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_MatrizDeRiesgosProfesionales & ", UserPermissionsProfile2);}"" /> Administración de Matriz de Riesgos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CargaDeMatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_CargaDeMatrizDeRiesgosProfesionales & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_CargaDeMatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_MatrizDeRiesgosProfesionales & ", UserPermissionsProfile2);}"" /> Carga de matriz de riesgos profesionales<BR />"
							End If
'							If aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_05_PERMISSIONS4 Then
'								Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions3b"" ID=""UserPermissions3bChPm"" VALUE=""" & N_05_PERMISSIONS4 & """"
'									If aUserComponent(L_PERMISSIONS4_USER) And N_05_PERMISSIONS4 Then
'										Response.Write " CHECKED=""1"""
'									End If
'								Response.Write " /> Fondo de ahorro capitalizable (FONAC)<BR />"
'							End If
'							If aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_06_PERMISSIONS4 Then
'								Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_06_PERMISSIONS4 & """"
'									If aUserComponent(L_PERMISSIONS4_USER) And N_06_PERMISSIONS4 Then
'										Response.Write " CHECKED=""1"""
'									End If
'								Response.Write " /> Sistema de ahorro para el retiro<BR />"
'							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_Catalogos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_Catalogos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_Catalogos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Catálogos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_02_AdministrarVentanillaUnica & ", UserPermissionsProfile2);} else {SetCheckboxesValue(" & N_02_AdministrarVentanillaUnica & ", UserPermissionsProfile2);}"" /> Ventanilla única<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AdministrarVentanillaUnica & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_02_VentanillaUnica & ", UserPermissionsProfile2);}"" /> Administrar Ventanilla única<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeBolsaDeTrabajoEscalafon & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_RegistroDeBolsaDeTrabajoEscalafon & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_RegistroDeBolsaDeTrabajoEscalafon & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Registro de bolsa de trabajo y escalafón<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AcumuladosAnuales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile2"" ID=""UserPermissionsProfile2ChPm"" VALUE=""" & N_02_AcumuladosAnuales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_02_AcumuladosAnuales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Acumulados anuales<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section03Div"" ID=""Section03Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_AdministracionDePlazas & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_AdministracionDePlazas & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_AdministracionDePlazas & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_03_ModificacionDePlazas & ", UserPermissionsProfile3);} else {SetCheckboxesValue(" & N_03_ModificacionDePlazas & ", UserPermissionsProfile3);}"" /> Administración de plazas<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_ModificacionDePlazas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_AdministracionDePlazas & ", UserPermissionsProfile3);}"" /> Modificación de plazas<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_EstructurasOcupacionales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_EstructurasOcupacionales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_EstructurasOcupacionales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_03_Catalogos & ", UserPermissionsProfile3); UncheckCheckboxesValue(" & N_03_ConsultaDeTabuladores & ", UserPermissionsProfile3); UncheckCheckboxesValue(" & N_03_RegistroDeTabuladores & ", UserPermissionsProfile3); UncheckCheckboxesValue(" & N_03_CargaUNIMED & ", UserPermissionsProfile3); } else { SetCheckboxesValue(" & N_03_Catalogos & ", UserPermissionsProfile3); SetCheckboxesValue(" & N_03_ConsultaDeTabuladores & ", UserPermissionsProfile3); SetCheckboxesValue(" & N_03_RegistroDeTabuladores & ", UserPermissionsProfile3); SetCheckboxesValue(" & N_03_CargaUNIMED & ", UserPermissionsProfile3); }"" /> Estructuras ocupacionales<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Catalogos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_Catalogos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_Catalogos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_EstructurasOcupacionales & ", UserPermissionsProfile3);}"" /> Catálogos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ConsultaDeTabuladores & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_ConsultaDeTabuladores & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_ConsultaDeTabuladores & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_EstructurasOcupacionales & ", UserPermissionsProfile3);}"" /> Consulta de tabuladores<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_RegistroDeTabuladores & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_RegistroDeTabuladores & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_RegistroDeTabuladores & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_EstructurasOcupacionales & ", UserPermissionsProfile3);}"" /> Registro de tabuladores<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_CargaUNIMED & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_CargaUNIMED & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_CargaUNIMED & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_EstructurasOcupacionales & ", UserPermissionsProfile3);}"" /> Carga UNIMED<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_SeleccionDePersonal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_SeleccionDePersonal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_SeleccionDePersonal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_03_BolsaDeTrabajo & ", UserPermissionsProfile3); UncheckCheckboxesValue(" & N_03_Escalafon & ", UserPermissionsProfile3); } else { SetCheckboxesValue(" & N_03_BolsaDeTrabajo & ", UserPermissionsProfile3); SetCheckboxesValue(" & N_03_Escalafon & ", UserPermissionsProfile3); }"" /> Selección de personal<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_BolsaDeTrabajo & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_BolsaDeTrabajo & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_BolsaDeTrabajo & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_SeleccionDePersonal & ", UserPermissionsProfile3);}"" /> Bolsa de trabajo<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Escalafon & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_Escalafon & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_Escalafon & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_03_SeleccionDePersonal & ", UserPermissionsProfile3);}"" /> Escalafón<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_DesarrolloHumano & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_DesarrolloHumano & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_DesarrolloHumano & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Desarrollo humano<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_PlaneacionDeRecursosHumanos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_PlaneacionDeRecursosHumanos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_PlaneacionDeRecursosHumanos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Planeación de recursos humanos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_BusquedaDeCentrosDeTrabajoYCentrosDePago & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_BusquedaDeCentrosDeTrabajoYCentrosDePago & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_BusquedaDeCentrosDeTrabajoYCentrosDePago & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Búsqueda de centros de trabajo y centros de pago<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Ventanilla única<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile3"" ID=""UserPermissionsProfile3ChPm"" VALUE=""" & N_03_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_03_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section04Div"" ID=""Section04Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_ConceptosDePago & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_ConceptosDePago & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_ConceptosDePago & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Conceptos de pago<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Empleados & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_Empleados & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_Empleados & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_04_RevisionDeNominas & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_C9_DevolucionesNoGravables & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_72_OtrasDeducciones & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_RegistroDeReclamos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_RevisionDeNominasReclamos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile4); } else { SetCheckboxesValue(" & N_04_RevisionDeNominas & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_C9_DevolucionesNoGravables & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_72_OtrasDeducciones & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_RegistroDeReclamos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_RevisionDeNominasReclamos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile4); }"" /> Empleados<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RevisionDeNominas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_RevisionDeNominas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_RevisionDeNominas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4);}"" /> Revisión de nóminas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_ReclamosDePagoPorAjustesYDeducciones & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_04_C9_DevolucionesNoGravables & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_72_OtrasDeducciones & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_RegistroDeReclamos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_RevisionDeNominasReclamos & ", UserPermissionsProfile4); UncheckCheckboxesValue(" & N_04_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile4); } else {SetCheckboxesValue(" & N_04_C9_DevolucionesNoGravables & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_72_OtrasDeducciones & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_RegistroDeReclamos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_RevisionDeNominasReclamos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_BajaDeRegistrosVigentesReclamos & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4);}"" /> Reclamos de pago por ajustes y deducciones<BR />"
							End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_C9_DevolucionesNoGravables & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> C9. Devoluciones no gravables<BR />"
								End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_71_DeduccionPorCobroDeSueldosIndebidos & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> 71. Deducción por cobro de sueldos indebidos<BR />"
								End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_72_OtrasDeducciones & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_72_OtrasDeducciones & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_72_OtrasDeducciones & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> 72. Otras deducciones<BR />"
								End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_RegistroDeReclamos & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> Registro de reclamos<BR />"
								End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RevisionDeNominasReclamos & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_RevisionDeNominasReclamos & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_RevisionDeNominasReclamos & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> Revisión de nóminas<BR />"
								End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_BajaDeRegistrosVigentesReclamos & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Empleados & ", UserPermissionsProfile4); SetCheckboxesValue(" & N_04_ReclamosDePagoPorAjustesYDeducciones & ", UserPermissionsProfile4);}"" /> Baja de registros vigentes<BR />"
								End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_CrearUnaNuevaNomina & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_CrearUnaNuevaNomina & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_CrearUnaNuevaNomina & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Crear una nueva nómina<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Prenomina & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_Prenomina & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_Prenomina & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Prenómina<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_CerrarNomina & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_CerrarNomina & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_CerrarNomina & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Cerrar nómina<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_NominasEspeciales & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_NominasEspeciales & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_NominasEspeciales & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Nóminas especiales<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Cheques & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_Cheques & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_Cheques & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Cheques<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_AperturaYCierreDeRegistros & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_AperturaYCierreDeRegistros & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_AperturaYCierreDeRegistros & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Apertura y cierre de registros<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Catalogos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_Catalogos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_Catalogos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_04_UsuariosDelSistema & ", UserPermissionsProfile4);} else {SetCheckboxesValue(" & N_04_UsuariosDelSistema & ", UserPermissionsProfile4);}"" /> Catálogos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_UsuariosDelSistema & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_UsuariosDelSistema & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_UsuariosDelSistema & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_04_Catalogos & ", UserPermissionsProfile4);}"" /> Usuarios del sistema<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Ventanilla única<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_TableroDeControlDeProcesos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_TableroDeControlDeProcesos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_TableroDeControlDeProcesos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Tablero de control de procesos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_EjercicioBimestralDelSAR & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile4"" ID=""UserPermissionsProfile4ChPm"" VALUE=""" & N_04_EjercicioBimestralDelSAR & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_04_EjercicioBimestralDelSAR & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Ejercicio bimestral del SAR<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section05Div"" ID=""Section05Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_EstructurasProgramaticas & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_EstructurasProgramaticas & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_EstructurasProgramaticas & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Estructuras programáticas<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ClasificadorPorObjetoDelGasto & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_ClasificadorPorObjetoDelGasto & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_ClasificadorPorObjetoDelGasto & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Clasificador por objeto del gasto<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_AdministracionDelPresupuesto & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_AdministracionDelPresupuesto & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_AdministracionDelPresupuesto & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administración del presupuesto<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ConsultaDePresupuesto & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_ConsultaDePresupuesto & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_ConsultaDePresupuesto & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Consulta de presupuesto<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_CosteoDePlazas & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_CosteoDePlazas & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_CosteoDePlazas & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Costeo de plazas<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ReportesSobreElCosteoDePlazas & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_ReportesSobreElCosteoDePlazas & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_ReportesSobreElCosteoDePlazas & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes sobre el costeo de plazas<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_RegistroDeUnCosteoComoOriginal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_RegistroDeUnCosteoComoOriginal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_RegistroDeUnCosteoComoOriginal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Registro de un costeo como original<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_Catalogos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_Catalogos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_Catalogos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Catálogos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Ventanilla única<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile5"" ID=""UserPermissionsProfile5ChPm"" VALUE=""" & N_05_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section06Div"" ID=""Section06Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_EmisionDeLicenciasPorComisionSindical & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_EmisionDeLicenciasPorComisionSindical & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_EmisionDeLicenciasPorComisionSindical & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Emisión de licencias por comisión sindical<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_Reportes & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_Reportes & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_Reportes & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_TableroDeDontrolDeProcesos & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_TableroDeDontrolDeProcesos & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_TableroDeDontrolDeProcesos & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Tablero de control de procesos<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_06_AdministrarVentanillaUnica & ", UserPermissionsProfile6); } else { SetCheckboxesValue(" & N_06_AdministrarVentanillaUnica & ", UserPermissionsProfile6); }"" /> Ventanilla única<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_AdministrarVentanillaUnica & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_06_VentanillaUnica & ", UserPermissionsProfile6);}"" /> Administrar Ventanilla única<BR />"
							End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile6"" ID=""UserPermissionsProfile6ChPm"" VALUE=""" & N_06_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_06_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section07Div"" ID=""Section07Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Personal & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_Personal & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_Personal & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_07_AsignacionDeNumeroTemporalDeEmpleado & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_AdministracionDePersonal & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ConsultaDePersonal & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ConsultaDePlazas & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ReportesPersonal & ", UserPermissionsProfile7);} else {SetCheckboxesValue(" & N_07_AsignacionDeNumeroTemporalDeEmpleado & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_AdministracionDePersonal & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ConsultaDePersonal & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ConsultaDePlazas & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ReportesPersonal & ", UserPermissionsProfile7);}"" /> Personal<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_AsignacionDeNumeroTemporalDeEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_AsignacionDeNumeroTemporalDeEmpleado & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_AsignacionDeNumeroTemporalDeEmpleado & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Personal & ", UserPermissionsProfile7);}"" /> Asignación de número temporal de empleado<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_AdministracionDePersonal & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_AdministracionDePersonal & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_AdministracionDePersonal & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Personal & ", UserPermissionsProfile7);}"" /> Administración de personal<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ConsultaDePersonal & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ConsultaDePersonal & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ConsultaDePersonal & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Personal & ", UserPermissionsProfile7);}"" /> Consulta de personal<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ConsultaDePlazas & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ConsultaDePlazas & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ConsultaDePlazas & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Personal & ", UserPermissionsProfile7);}"" /> Consulta de plazas<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesPersonal & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ReportesPersonal & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ReportesPersonal & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Personal & ", UserPermissionsProfile7);}"" /> Reportes<BR />"
							End If

						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Prestaciones & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_Prestaciones & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_Prestaciones & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_07_PrestacionesEIncidencias & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_EntregasDeHojasUnicasDeServicio & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ReportesPrestaciones & ", UserPermissionsProfile7);} else {SetCheckboxesValue(" & N_07_PrestacionesEIncidencias & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_EntregasDeHojasUnicasDeServicio & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ReportesPrestaciones & ", UserPermissionsProfile7);}"" /> Prestaciones<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_PrestacionesEIncidencias & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Prestaciones & ", UserPermissionsProfile7);}"" /> Prestaciones e incidencias<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_EntregasDeHojasUnicasDeServicio & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Prestaciones & ", UserPermissionsProfile7);}"" /> Entregas de hojas únicas de servicio<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesPrestaciones & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ReportesPrestaciones & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ReportesPrestaciones & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Prestaciones & ", UserPermissionsProfile7);}"" /> Reportes<BR />"
							End If

						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Informatica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_Informatica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_Informatica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) {UncheckCheckboxesValue(" & N_07_Empleados & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ChequesYDepositos & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_ReportesInformatica & ", UserPermissionsProfile7); UncheckCheckboxesValue(" & N_07_RevisionDeNominas & ", UserPermissionsProfile7);} else {SetCheckboxesValue(" & N_07_Empleados & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ChequesYDepositos & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_ReportesInformatica & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_RevisionDeNominas & ", UserPermissionsProfile7);}"" /> Informática<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Empleados & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_Empleados & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_Empleados & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Informatica & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_RevisionDeNominas & ", UserPermissionsProfile7);} else {UncheckCheckboxesValue(" & N_07_RevisionDeNominas & ", UserPermissionsProfile7);} "" /> Empleados<BR />"
							End If
								If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_RevisionDeNominas & ",", vbBinaryCompare) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_RevisionDeNominas & """"
										If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_RevisionDeNominas & ",", vbBinaryCompare) > 0 Then
											Response.Write " CHECKED=""1"""
										End If
									Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Informatica & ", UserPermissionsProfile7); SetCheckboxesValue(" & N_07_Empleados & ", UserPermissionsProfile7);}"" /> Revisión de nóminas<BR />"
								End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ChequesYDepositos & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ChequesYDepositos & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ChequesYDepositos & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Informatica & ", UserPermissionsProfile7);}"" /> Cheques y depósitos<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesInformatica & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ReportesInformatica & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ReportesInformatica & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_07_Informatica & ", UserPermissionsProfile7);}"" /> Reportes<BR />"
							End If

						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Presupuesto & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_Presupuesto & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_Presupuesto & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Presupuesto<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_TableroDeControl & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_TableroDeControl & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_TableroDeControl & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Tablero de control<BR />"
						End If
						'If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
						'	Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions"" ID=""UserPermissionsChPm"" VALUE=""" & N_TOOLS_PERMISSIONS & """"
						'		If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_05_CosteoDePlazas & ",", vbBinaryCompare) > 0 Then
						'			Response.Write " CHECKED=""1"""
						'		End If
						'	Response.Write " /> Herramientas administrativas<BR />"
						'End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesGuardados & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_ReportesGuardados & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_ReportesGuardados & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Reportes guardados<BR />"
						End If
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_AdministrarNormateca & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile7"" ID=""UserPermissionsProfile7ChPm"" VALUE=""" & N_07_AdministrarNormateca & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_07_AdministrarNormateca & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " /> Administrar normateca<BR />"
						End If
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""Section08Div"" ID=""Section08Div"" STYLE=""display: none"">"
						If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
							Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile8"" ID=""UserPermissionsProfile8ChPm"" VALUE=""" & N_08_VentanillaUnica & """"
								If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_08_VentanillaUnica & ",", vbBinaryCompare) > 0 Then
									Response.Write " CHECKED=""1"""
								End If
							Response.Write " onClick=""if (!this.checked) { UncheckCheckboxesValue(" & N_08_AdministrarVentanillaUnica & ", UserPermissionsProfile8); UncheckCheckboxesValue(" & N_08_ModificarResponsableDeDocumento & ", UserPermissionsProfile8); } else { SetCheckboxesValue(" & N_08_AdministrarVentanillaUnica & ", UserPermissionsProfile8); SetCheckboxesValue(" & N_08_ModificarResponsableDeDocumento & ", UserPermissionsProfile8); }"" /> Ventanilla única<BR />"
						End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile8"" ID=""UserPermissionsProfile8ChPm"" VALUE=""" & N_08_AdministrarVentanillaUnica & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_08_VentanillaUnica & ", UserPermissionsProfile8);}"" /> Administrar Ventanilla única<BR />"
							End If
							If (StrComp("0", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) = 0) Or (StrComp(aLoginComponent(N_USER_PERMISSIONS2_LOGIN), "-1", vbBinaryCompare) = 0) Or (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_ModificarResponsableDeDocumento & ",", vbBinaryCompare) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissionsProfile8"" ID=""UserPermissionsProfile8ChPm"" VALUE=""" & N_08_ModificarResponsableDeDocumento & """"
									If InStr(1, "," & aUserComponent(L_PERMISSIONS2_USER) & ",", "," & N_08_ModificarResponsableDeDocumento & ",", vbBinaryCompare) > 0 Then
										Response.Write " CHECKED=""1"""
									End If
								Response.Write " onClick=""if (this.checked) {SetCheckboxesValue(" & N_08_VentanillaUnica & ", UserPermissionsProfile8);}"" /> Modificar Responsable de Documento<BR />"
							End If
					Response.Write "</DIV>"

					If (aUserComponent(N_ID_USER) > 0) And (aUserComponent(N_PROFILE_ID_USER) > 0) Then
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "ShowPermissionsForProfile(document.UserFrm.ProfileID.value);" & vbNewLine
							'Response.Write "SendURLValuesToForm('UserPermissionsProfile" & Right(("00" & aUserComponent(N_PROFILE_ID_USER)), Len("00")) & "=' + " & aLoginComponent(N_USER_PERMISSIONS4_LOGIN) & ", document.UserFrm);" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					End If

					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_FILES_PERMISSIONS Then
						Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""UserPermissions"" ID=""UserPermissionsChPm"" VALUE=""" & N_DELETE_FILES_PERMISSIONS & """"
							If aUserComponent(L_PERMISSIONS_USER) And N_DELETE_FILES_PERMISSIONS Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " /> Borrar archivos<BR />"
					End If

					Response.Write "<BR />"
				Response.Write "</FONT>"
			Response.Write "</DIV>"
			Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""770"" HEIGHT=""1"" /><BR />"

			Response.Write "<BR />"
			If aUserComponent(N_ID_USER) = -2 Then
				If (Not B_PORTAL) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If (Not B_PORTAL) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveUserWngDiv']); UserFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Users'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveUserWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	Set oRecordset = Nothing
	DisplayUserForm = lErrorNumber
	Err.Clear
End Function

Function DisplayUserAsHiddenFields(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an user using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aUserComponent
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayUserAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aUserComponent(N_ID_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserAccessKey"" ID=""UserAccessKeyHdn"" VALUE=""" & aUserComponent(S_ACCESS_KEY_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPassword"" ID=""UserPasswordHdn"" VALUE=""" & aUserComponent(S_PASSWORD_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserName"" ID=""UserNameHdn"" VALUE=""" & aUserComponent(S_NAME_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserLastName"" ID=""UserLastNameHdn"" VALUE=""" & aUserComponent(S_LAST_NAME_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserEmail"" ID=""UserEmailHdn"" VALUE=""" & aUserComponent(S_EMAIL_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions"" ID=""UserPermissionsHdn"" VALUE=""" & aUserComponent(L_PERMISSIONS_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions2"" ID=""UserPermissions2Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS2_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions3"" ID=""UserPermissions3Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS3_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserPermissions4"" ID=""UserPermissions4Hdn"" VALUE=""" & aUserComponent(L_PERMISSIONS4_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports"" ID=""PermissionReportsHdn"" VALUE=""" & aUserComponent(L_PERMISSION_REPORTS_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports2"" ID=""PermissionReports2Hdn"" VALUE=""" & aUserComponent(L_PERMISSION_REPORTS2_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionAreaID"" ID=""PermissionAreaIDHdn"" VALUE=""" & aUserComponent(S_PERMISSIONS_AREAS_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionZoneID"" ID=""PermissionZoneIDHdn"" VALUE=""" & aUserComponent(L_PERMISSIONS_ZONE_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BossEmail"" ID=""BossEmailHdn"" VALUE=""" & aUserComponent(S_BOSS_EMAIL_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AdditionalEmail"" ID=""AdditionalEmailHdn"" VALUE=""" & aUserComponent(S_ADDITIONAL_EMAIL_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileID"" ID=""ProfileIDHdn"" VALUE=""" & aUserComponent(N_PROFILE_ID_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserActive"" ID=""UserActiveHdn"" VALUE=""" & aUserComponent(N_ACTIVE_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserBlocked"" ID=""UserBlockedHdn"" VALUE=""" & aUserComponent(N_BLOCKED_USER) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TechSupport"" ID=""TechSupportHdn"" VALUE=""" & aUserComponent(N_TECH_SUPPORT_USER) & """ />"

	DisplayUserAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayUsersTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the users from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aUserComponent
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayUsersTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber

	lErrorNumber = GetUsers(oRequest, oADODBConnection, aUserComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""360"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,Usuario,Clave,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,150,70,100", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Usuario,Clave", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,220,120", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("UserID").Value), oRequest("UserID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If CInt(oRecordset.Fields("SecurityLock").Value) >= CInt(GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION)) Then
						sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=Users&UserToUnlock=" & CleanStringForJavaScript(CStr(oRecordset.Fields("UserAccessKey").Value)) & "&Unlock=1""><IMG SRC=""Images/IcnLock.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Usuario bloqueado. Para reactivar su cuenta presione aquí."" BORDER=""0""/></A>"
					End If
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""UserID"" ID=""UserIDRd"" VALUE=""" & CStr(oRecordset.Fields("UserID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""UserID"" ID=""UserIDChk"" VALUE=""" & CStr(oRecordset.Fields("UserID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT SIZE=""1"""
						sRowContents = sRowContents & " TITLE=""e-mail: " & CleanStringForHTML(CStr(oRecordset.Fields("UserEmail").Value)) & "&#13;e-mail del jefe directo: " & CleanStringForHTML(CStr(oRecordset.Fields("BossEmail").Value)) & "&#13;e-mail adicional: " & CleanStringForHTML(CStr(oRecordset.Fields("AdditionalEmail").Value)) & """"
						If CInt(oRecordset.Fields("UserActive").Value) = 0 Then sRowContents = sRowContents & " COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """"
					sRowContents = sRowContents & ">" & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("UserLastName").Value) & ", " & CStr(oRecordset.Fields("UserName").Value)) & sBoldEnd & "</FONT>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT SIZE=""1"""
						If CInt(oRecordset.Fields("UserActive").Value) = 0 Then sRowContents = sRowContents & " COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """"
					sRowContents = sRowContents & ">" & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("UserAccessKey").Value)) & sBoldEnd & "</FONT>"
					If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Users&UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If B_USE_SMTP And (Not B_PORTAL) Then
								sRowContents = sRowContents & "<A HREF=""SendUserEmail.asp?UserID=" & CStr(oRecordset.Fields("UserID").Value) & """ TARGET=""SendUserEmail"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnEmail.gif"" WIDTH=""11"" HEIGHT=""8"" ALT=""Enviar Contraseña"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (Not B_PORTAL) Then
								If aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value) Then
									sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />&nbsp;&nbsp;&nbsp;"
								Else
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Users&UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&Delete=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								End If
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								If CInt(oRecordset.Fields("UserActive").Value) = 0 Then 
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Users&UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""El usuario tendrá acceso a su cuenta corporativa de correo electrónico"" BORDER=""0"" /></A>"
								Else
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Users&UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""El usuario no tendrá acceso a su cuenta corporativa de correo electrónico"" BORDER=""0"" /></A>"
								End If
							End If
						sRowContents = sRowContents & "&nbsp;"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen usuarios registrados en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayUsersTable = lErrorNumber
	Err.Clear
End Function
%>