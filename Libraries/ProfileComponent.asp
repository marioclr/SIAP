<%
Const N_ID_PROFILE = 0
Const S_NAME_PROFILE = 1
Const N_PERMISSIONS_PROFILE = 2
Const N_PERMISSIONS2_PROFILE = 3
Const N_PERMISSIONS3_PROFILE = 4
Const N_PERMISSIONS4_PROFILE = 5
Const N_PERMISSION_REPORTS_PROFILE = 6
Const N_PERMISSION_REPORTS2_PROFILE = 7
Const B_CHECK_FOR_DUPLICATED_PROFILE = 8
Const B_IS_DUPLICATED_PROFILE = 9
Const B_COMPONENT_INITIALIZED_PROFILE = 10

Const N_PROFILE_COMPONENT_SIZE = 10

Dim aProfileComponent()
Redim aProfileComponent(N_PROFILE_COMPONENT_SIZE)

Function InitializeProfileComponent(oRequest, aProfileComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Profile Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aProfileComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeProfileComponent"
	Dim iItem
	Redim Preserve aProfileComponent(N_PROFILE_COMPONENT_SIZE)

	If IsEmpty(aProfileComponent(N_ID_PROFILE)) Then
		If Len(oRequest("ProfileID").Item) > 0 Then
			aProfileComponent(N_ID_PROFILE) = CLng(oRequest("ProfileID").Item)
		Else
			aProfileComponent(N_ID_PROFILE) = -1
		End If
	End If

	If IsEmpty(aProfileComponent(S_NAME_PROFILE)) Then
		If Len(oRequest("ProfileName").Item) > 0 Then
			aProfileComponent(S_NAME_PROFILE) = oRequest("ProfileName").Item
		Else
			aProfileComponent(S_NAME_PROFILE) = ""
		End If
	End If
	aProfileComponent(S_NAME_PROFILE) = Left(aProfileComponent(S_NAME_PROFILE), 100)

	If IsEmpty(aProfileComponent(N_PERMISSIONS_PROFILE)) Then
		If Len(oRequest("ProfilePermissions").Item) > 0 Then
			If InStr(1, oRequest("ProfilePermissions").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSIONS_PROFILE) = 0
				For Each iItem In oRequest("ProfilePermissions")
					aProfileComponent(N_PERMISSIONS_PROFILE) = aProfileComponent(N_PERMISSIONS_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSIONS_PROFILE) = CLng(oRequest("ProfilePermissions").Item)
			End If
		Else
			aProfileComponent(N_PERMISSIONS_PROFILE) = 0
		End If
	End If

	If IsEmpty(aProfileComponent(N_PERMISSIONS2_PROFILE)) Then
		If Len(oRequest("ProfilePermissions2").Item) > 0 Then
			If InStr(1, oRequest("ProfilePermissions2").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSIONS2_PROFILE) = 0
				For Each iItem In oRequest("ProfilePermissions2")
					aProfileComponent(N_PERMISSIONS2_PROFILE) = aProfileComponent(N_PERMISSIONS2_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSIONS2_PROFILE) = CLng(oRequest("ProfilePermissions2").Item)
			End If
		Else
			aProfileComponent(N_PERMISSIONS2_PROFILE) = 0
		End If
	End If

	If IsEmpty(aProfileComponent(N_PERMISSIONS3_PROFILE)) Then
		If Len(oRequest("ProfilePermissions3").Item) > 0 Then
			If InStr(1, oRequest("ProfilePermissions3").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSIONS3_PROFILE) = 0
				For Each iItem In oRequest("ProfilePermissions3")
					aProfileComponent(N_PERMISSIONS3_PROFILE) = aProfileComponent(N_PERMISSIONS3_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSIONS3_PROFILE) = CLng(oRequest("ProfilePermissions3").Item)
			End If
		Else
			aProfileComponent(N_PERMISSIONS3_PROFILE) = 0
		End If
	End If

	If IsEmpty(aProfileComponent(N_PERMISSIONS4_PROFILE)) Then
		If Len(oRequest("ProfilePermissions4").Item) > 0 Then
			If InStr(1, oRequest("ProfilePermissions4").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSIONS4_PROFILE) = 0
				For Each iItem In oRequest("ProfilePermissions4")
					aProfileComponent(N_PERMISSIONS4_PROFILE) = aProfileComponent(N_PERMISSIONS4_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSIONS4_PROFILE) = CLng(oRequest("ProfilePermissions4").Item)
			End If
		Else
			aProfileComponent(N_PERMISSIONS4_PROFILE) = 0
		End If
	End If

	If IsEmpty(aProfileComponent(N_PERMISSION_REPORTS_PROFILE)) Then
		If Len(oRequest("PermissionReports").Item) > 0 Then
			If InStr(1, oRequest("PermissionReports").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = 0
				For Each iItem In oRequest("PermissionReports")
					aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = aProfileComponent(N_PERMISSION_REPORTS_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = CLng(oRequest("PermissionReports").Item)
			End If
		Else
			aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = 0
		End If
	End If

	If IsEmpty(aProfileComponent(N_PERMISSION_REPORTS2_PROFILE)) Then
		If Len(oRequest("PermissionReports2").Item) > 0 Then
			If InStr(1, oRequest("PermissionReports2").Item, ",", vbBinaryCompare) > 1 Then
				aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = 0
				For Each iItem In oRequest("PermissionReports2")
					aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) + CLng(iItem)
				Next
			Else
				aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = CLng(oRequest("PermissionReports2").Item)
			End If
		Else
			aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = 0
		End If
	End If

	aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE) = True
	aProfileComponent(B_IS_DUPLICATED_PROFILE) = False

	aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE) = True
	InitializeProfileComponent = Err.number
	Err.Clear
End Function

Function AddProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new profile into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddProfile"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "UserProfiles", "ProfileID", "", 1, aProfileComponent(N_ID_PROFILE), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE) Then
			lErrorNumber = CheckExistencyOfProfile(aProfileComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aProfileComponent(B_IS_DUPLICATED_PROFILE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro con el nombre " & aProfileComponent(S_NAME_PROFILE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckProfileInformationConsistency(aProfileComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo guardar la información del nuevo registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into UserProfiles (ProfileID, ProfileName, ProfilePermissions, ProfilePermissions2, ProfilePermissions3, ProfilePermissions4, PermissionReports, PermissionReports2) Values (" & aProfileComponent(N_ID_PROFILE) & ", '" & Replace(aProfileComponent(S_NAME_PROFILE), "'", "") & "', " & aProfileComponent(N_PERMISSIONS_PROFILE) & ", " & aProfileComponent(N_PERMISSIONS2_PROFILE) & ", " & aProfileComponent(N_PERMISSIONS3_PROFILE) & ", " & aProfileComponent(N_PERMISSIONS4_PROFILE) & ", " & aProfileComponent(N_PERMISSION_REPORTS_PROFILE) & ", " & aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	AddProfile = lErrorNumber
	Err.Clear
End Function

Function GetProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a profile from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProfile"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From UserProfiles Where ProfileID=" & aProfileComponent(N_ID_PROFILE), "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aProfileComponent(S_NAME_PROFILE) = CStr(oRecordset.Fields("ProfileName").Value)
				aProfileComponent(N_PERMISSIONS_PROFILE) = CLng(oRecordset.Fields("ProfilePermissions").Value)
				aProfileComponent(N_PERMISSIONS2_PROFILE) = CLng(oRecordset.Fields("ProfilePermissions2").Value)
				aProfileComponent(N_PERMISSIONS3_PROFILE) = CLng(oRecordset.Fields("ProfilePermissions3").Value)
				aProfileComponent(N_PERMISSIONS4_PROFILE) = CLng(oRecordset.Fields("ProfilePermissions4").Value)
				aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = CLng(oRecordset.Fields("PermissionReports").Value)
				aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = CLng(oRecordset.Fields("PermissionReports2").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetProfile = lErrorNumber
	Err.Clear
End Function

Function GetProfiles(oRequest, oADODBConnection, aProfileComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the profiles from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfileComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProfiles"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From UserProfiles Where (ProfileID>-1) Order By ProfileID", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetProfiles = lErrorNumber
	Err.Clear
End Function

Function ModifyProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing profile in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyProfile"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE) Then
			lErrorNumber = CheckExistencyOfProfile(aProfileComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aProfileComponent(B_IS_DUPLICATED_PROFILE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro con el nombre " & aProfileComponent(S_NAME_PROFILE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckProfileInformationConsistency(aProfileComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update UserProfiles Set ProfileName='" & Replace(aProfileComponent(S_NAME_PROFILE), "'", "") & "', ProfilePermissions=" & aProfileComponent(N_PERMISSIONS_PROFILE) & ", ProfilePermissions2=" & aProfileComponent(N_PERMISSIONS2_PROFILE) & ", ProfilePermissions3=" & aProfileComponent(N_PERMISSIONS3_PROFILE) & ", ProfilePermissions4=" & aProfileComponent(N_PERMISSIONS4_PROFILE) & ", PermissionReports=" & aProfileComponent(N_PERMISSION_REPORTS_PROFILE) & ", PermissionReports2=" & aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) & " Where (ProfileID=" & aProfileComponent(N_ID_PROFILE) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set UserPermissions=" & aProfileComponent(N_PERMISSIONS_PROFILE) & ", UserPermissions2=" & aProfileComponent(N_PERMISSIONS2_PROFILE) & ", UserPermissions3=" & aProfileComponent(N_PERMISSIONS3_PROFILE) & ", UserPermissions4=" & aProfileComponent(N_PERMISSIONS4_PROFILE) & ", PermissionReports=" & aProfileComponent(N_PERMISSION_REPORTS_PROFILE) & ", PermissionReports2=" & aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) & " Where (ProfileID=" & aProfileComponent(N_ID_PROFILE) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	ModifyProfile = lErrorNumber
	Err.Clear
End Function

Function RemoveProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a profile from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveProfile"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From UserProfiles Where (ProfileID=" & aProfileComponent(N_ID_PROFILE) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sErrorDescription = "No se pudieron actualizar los siniestros que afectaron a la cobertura."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set ProfileID=-1 Where (ProfileID=" & aProfileComponent(N_ID_PROFILE) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveProfile = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfProfile(aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific profile exists in the database
'Inputs:  aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfProfile"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If Len(aProfileComponent(S_NAME_PROFILE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From UserProfiles Where (ProfileID<>" & aProfileComponent(N_ID_PROFILE) & ") And (ProfileName='" & Replace(aProfileComponent(S_NAME_PROFILE), "'", "") & "')", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aProfileComponent(B_IS_DUPLICATED_PROFILE) = True
				aProfileComponent(N_ID_PROFILE) = CLng(oRecordset.Fields("ProfileID").Value)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfProfile = lErrorNumber
	Err.Clear
End Function

Function CheckProfileInformationConsistency(aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aProfileComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckProfileInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aProfileComponent(N_ID_PROFILE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aProfileComponent(S_NAME_PROFILE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del registro está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aProfileComponent(N_PERMISSIONS_PROFILE)) Then aProfileComponent(N_PERMISSIONS_PROFILE) = 0
	If Not IsNumeric(aProfileComponent(N_PERMISSIONS2_PROFILE)) Then aProfileComponent(N_PERMISSIONS2_PROFILE) = 0
	If Not IsNumeric(aProfileComponent(N_PERMISSIONS3_PROFILE)) Then aProfileComponent(N_PERMISSIONS3_PROFILE) = 0
	If Not IsNumeric(aProfileComponent(N_PERMISSIONS4_PROFILE)) Then aProfileComponent(N_PERMISSIONS4_PROFILE) = 0
	If Not IsNumeric(aProfileComponent(N_PERMISSION_REPORTS_PROFILE)) Then aProfileComponent(N_PERMISSION_REPORTS_PROFILE) = 0
	If Not IsNumeric(aProfileComponent(N_PERMISSION_REPORTS2_PROFILE)) Then aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos:" & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckProfileInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayProfileForm(oRequest, oADODBConnection, sAction, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a profile from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfileForm"
	Dim sNames
	Dim sTempNames
	Dim lErrorNumber

	If aProfileComponent(N_ID_PROFILE) <> -1 Then
		lErrorNumber = GetProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckProfileFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.ProfileName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.ProfileName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckProfileFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ProfileFrm"" ID=""ProfileFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckProfileFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Profiles"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileID"" ID=""ProfileIDHdn"" VALUE=""" & aProfileComponent(N_ID_PROFILE) & """ />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre: </FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""ProfileName"" ID=""ProfileNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & aProfileComponent(S_NAME_PROFILE) & """ CLASS=""TextFields"" /><BR />"
			If B_ISSSTE Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsHdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS_PROFILE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions2"" ID=""ProfilePermissions2Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS2_PROFILE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions3"" ID=""ProfilePermissions3Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS3_PROFILE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions4"" ID=""ProfilePermissions4Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS4_PROFILE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports"" ID=""PermissionReportsHdn"" VALUE=""" & aProfileComponent(N_PERMISSION_REPORTS_PROFILE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports2"" ID=""PermissionReports2Hdn"" VALUE=""" & aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) & """ />"
			Else
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_ADD_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_ADD_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Agregar registros<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_MODIFY_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_MODIFY_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Modificar registros<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_REMOVE_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_REMOVE_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Eliminar registros<BR />"
					Response.Write "<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_BUDGET_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_BUDGET_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de presupuestos<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_AREAS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_AREAS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de áreas<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_POSITIONS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_POSITIONS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de puestos<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_JOBS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_JOBS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de plazas<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_EMPLOYEES_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_EMPLOYEES_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de empleados<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_EMPLOYEE_PAYROLL_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_EMPLOYEE_PAYROLL_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de pagos a empleados<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_SADE_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_SADE_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Desarrollo Humano<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_PAYROLL_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_PAYROLL_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de la nómina<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_PAYMENTS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_PAYMENTS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración de cheques<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_REPORTS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_REPORTS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Ver reportes<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_TOOLS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_TOOLS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Usar herramientas<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_CATALOGS_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_CATALOGS_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Administración catálogos<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_TACO_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_TACO_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Tablero de control<BR />"

					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsChPm"" VALUE=""" & N_DELETE_FILES_PERMISSIONS & """"
						If aProfileComponent(N_PERMISSIONS_PROFILE) And N_DELETE_FILES_PERMISSIONS Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " /> Borrar archivos digitalizados<BR />"

				Response.Write "<BR /></FONT>"
			End If
			Response.Write "<BR />"

			If aProfileComponent(N_ID_PROFILE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveUserWngDiv']); ProfileFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Profiles'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveCatalogWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayProfileForm = lErrorNumber
	Err.Clear
End Function

Function DisplayProfileAsHiddenFields(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a profile using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfileAsHiddenFields"

	If Len(aProfileComponent(S_URL_PROFILE)) > 0 Then Call DisplayURLParametersAsHiddenValues(aProfileComponent(S_URL_PROFILE))
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileID"" ID=""ProfileIDHdn"" VALUE=""" & aProfileComponent(N_ID_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileName"" ID=""ProfileNameHdn"" VALUE=""" & aProfileComponent(S_NAME_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions"" ID=""ProfilePermissionsHdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions2"" ID=""ProfilePermissions2Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS2_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions3"" ID=""ProfilePermissions3Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS3_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfilePermissions4"" ID=""ProfilePermissions4Hdn"" VALUE=""" & aProfileComponent(N_PERMISSIONS4_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports"" ID=""PermissionReportsHdn"" VALUE=""" & aProfileComponent(N_PERMISSION_REPORTS_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PermissionReports2"" ID=""PermissionReports2Hdn"" VALUE=""" & aProfileComponent(N_PERMISSION_REPORTS2_PROFILE) & """ />"

	DisplayProfileAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayProfilesTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the profiles from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfilesTable"
	Dim sNames
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

	lErrorNumber = GetProfiles(oRequest, oADODBConnection, aProfileComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,Nombre,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,451,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Nombre", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,451", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ProfileID").Value), oRequest("ProfileID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ProfileID"" ID=""ProfileIDRd"" VALUE=""" & CStr(oRecordset.Fields("ProfileID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ProfileID"" ID=""ProfileIDChk"" VALUE=""" & CStr(oRecordset.Fields("ProfileID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProfileName"))) & sBoldEnd
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("ProfileID").Value) <> 0 Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Profiles&ProfileID=" & CStr(oRecordset.Fields("ProfileID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_PERMISSIONS) = N_DELETE_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Profiles&ProfileID=" & CStr(oRecordset.Fields("ProfileID").Value) & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayProfilesTable = lErrorNumber
	Err.Clear
End Function
%>