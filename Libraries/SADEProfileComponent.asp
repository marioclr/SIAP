<%
Const N_ID_PROFILE = 0
Const N_PARENT_ID_PROFILE = 1
Const S_NAME_PROFILE = 2
Const S_COMMENTS_PROFILE = 3
Const B_CHECK_FOR_DUPLICATED_PROFILE = 4
Const B_MODIFY_PROFILE = 5
Const S_QUERY_CONDITION_PROFILE = 6
Const S_COURSES_PROFILE = 7
Const S_COURSES_START_DATES_PROFILE = 8
Const S_COURSES_LAST_DATES_PROFILE = 9
Const S_USERS_PROFILE = 10
Const B_IS_DUPLICATED_PROFILE = 11
Const S_TARGET_PAGE_PROFILE = 12
Const N_ID_SELECTED_PROFILE = 13
Const B_COMPONENT_INITIALIZED_PROFILE = 14

Const N_PROFILE_COMPONENT_SIZE = 14

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
	Dim aUserIdentifier
	Redim Preserve aProfileComponent(N_PROFILE_COMPONENT_SIZE)

	If IsEmpty(aProfileComponent(N_ID_PROFILE)) Then
		If Len(oRequest("ProfileID").Item) > 0 Then
			aProfileComponent(N_ID_PROFILE) = CLng(oRequest("ProfileID").Item)
		Else
			aProfileComponent(N_ID_PROFILE) = -1
		End If
	End If
	If aProfileComponent(N_ID_PROFILE) = -1 Then
		aUserIdentifier = Split(oRequest("UserIdentifier").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
		aProfileComponent(N_ID_PROFILE) = CLng(aUserIdentifier(1))
	End If

	If IsEmpty(aProfileComponent(N_PARENT_ID_PROFILE)) Then
		If Len(oRequest("ParentID").Item) > 0 Then
			aProfileComponent(N_PARENT_ID_PROFILE) = CLng(oRequest("ParentID").Item)
		Else
			aProfileComponent(N_PARENT_ID_PROFILE) = -1
		End If
	End If

	If IsEmpty(aProfileComponent(S_NAME_PROFILE)) Then
		If Len(oRequest("ProfileName").Item) > 0 Then
			aProfileComponent(S_NAME_PROFILE) = oRequest("ProfileName").Item
		Else
			aProfileComponent(S_NAME_PROFILE) = ""
		End If
	End If
	aProfileComponent(S_NAME_PROFILE) = Left(aProfileComponent(S_NAME_PROFILE), 255)

	If IsEmpty(aProfileComponent(S_COMMENTS_PROFILE)) Then
		If Len(oRequest("ProfileComments").Item) > 0 Then
			aProfileComponent(S_COMMENTS_PROFILE) = oRequest("ProfileComments").Item
		Else
			aProfileComponent(S_COMMENTS_PROFILE) = ""
		End If
	End If
	aProfileComponent(S_COMMENTS_PROFILE) = Left(aProfileComponent(S_COMMENTS_PROFILE), 255)

	If IsEmpty(aProfileComponent(S_COURSES_PROFILE)) Then
		If Len(oRequest("ProfileCourses").Item) > 0 Then
			aProfileComponent(S_COURSES_PROFILE) = oRequest("ProfileCourses").Item
		Else
			aProfileComponent(S_COURSES_PROFILE) = ""
		End If
	End If

	If IsEmpty(aProfileComponent(S_COURSES_START_DATES_PROFILE)) Then
		If Len(oRequest("ProfileCoursesStartDates").Item) > 0 Then
			aProfileComponent(S_COURSES_START_DATES_PROFILE) = oRequest("ProfileCoursesStartDates").Item
		Else
			aProfileComponent(S_COURSES_START_DATES_PROFILE) = ""
		End If
	End If

	If IsEmpty(aProfileComponent(S_COURSES_LAST_DATES_PROFILE)) Then
		If Len(oRequest("ProfileCoursesLastDates").Item) > 0 Then
			aProfileComponent(S_COURSES_LAST_DATES_PROFILE) = oRequest("ProfileCoursesLastDates").Item
		Else
			aProfileComponent(S_COURSES_LAST_DATES_PROFILE) = ""
		End If
	End If

	If IsEmpty(aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE)) Then
		aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE) = (Len(oRequest("CheckDuplicatedProfiles").Item) > 0)
	End If
	aProfileComponent(B_IS_DUPLICATED_PROFILE) = False

	If IsEmpty(aProfileComponent(B_MODIFY_PROFILE)) Then
		aProfileComponent(B_MODIFY_PROFILE) = (Len(oRequest("ModifyProfile").Item) > 0)
	End If

	If IsEmpty(aProfileComponent(S_QUERY_CONDITION_PROFILE)) Then
		If Len(oRequest("ProfileCondition").Item) > 0 Then
			aProfileComponent(S_QUERY_CONDITION_PROFILE) = oRequest("ProfileCondition").Item
		Else
			aProfileComponent(S_QUERY_CONDITION_PROFILE) = ""
		End If
	End If

	If IsEmpty(aProfileComponent(S_TARGET_PAGE_PROFILE)) Then
		If Len(oRequest("ProfileTargetPage").Item) > 0 Then
			aProfileComponent(S_TARGET_PAGE_PROFILE) = oRequest("ProfileTargetPage").Item
		Else
			aProfileComponent(S_TARGET_PAGE_PROFILE) = GetASPFileName("")
		End If
	End If

	If IsEmpty(aProfileComponent(N_ID_SELECTED_PROFILE)) Then
		If Len(oRequest("SelectedProfileID").Item) > 0 Then
			aProfileComponent(N_ID_SELECTED_PROFILE) = CLng(oRequest("SelectedProfileID").Item)
		Else
			aProfileComponent(N_ID_SELECTED_PROFILE) = -1
		End If
	End If

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
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo curso."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, SADE_PREFIX & "Perfiles", "ID_Perfil", "", 1, aProfileComponent(N_ID_PROFILE), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aProfileComponent(B_CHECK_FOR_DUPLICATED_PROFILE) Then
			lErrorNumber = CheckExistencyOfProfile(aProfileComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If Not aProfileComponent(B_IS_DUPLICATED_PROFILE) Then
				If Not CheckProfileInformationConsistency(aProfileComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo guardar la información del nuevo curso."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "Perfiles (ID_Perfil, ID_Padre, Nombre_Perfil, Comentarios) Values (" & aProfileComponent(N_ID_PROFILE) & ", " & aProfileComponent(N_PARENT_ID_PROFILE) & ", '" & Replace(aProfileComponent(S_NAME_PROFILE), "'", "") & "', '" & Replace(aProfileComponent(S_COMMENTS_PROFILE), "'", "") & "')", "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
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
		sErrorDescription = "No se especificó el identificador del curso para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del curso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Perfiles Where ID_Perfil=" & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El curso especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aProfileComponent(N_PARENT_ID_PROFILE) = CLng(oRecordset.Fields("ID_Padre").Value)
				aProfileComponent(S_NAME_PROFILE) = CStr(oRecordset.Fields("Nombre_Perfil").Value)
				aProfileComponent(S_COMMENTS_PROFILE) = CStr(oRecordset.Fields("Comentarios").Value)
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
'Inputs:  oRequest, oADODBConnection, aProfileComponent
'Outputs: aProfileComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProfiles"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	sCondition = ""
	If Len(aProfileComponent(S_QUERY_CONDITION_PROFILE)) > 0 Then
		sCondition = " Where " & aProfileComponent(S_QUERY_CONDITION_PROFILE)
	End If
	sErrorDescription = "No se pudo obtener la información de los cursos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Perfiles " & sCondition & " Order By Nombre_Perfil", "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

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
		sErrorDescription = "No se especificó el identificador del curso a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckProfileInformationConsistency(aProfileComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "Perfiles Set ID_Padre=" & aProfileComponent(N_PARENT_ID_PROFILE) & ", Nombre_Perfil='" & Replace(aProfileComponent(S_NAME_PROFILE), "'", "") & "', Comentarios='" & Replace(aProfileComponent(S_COMMENTS_PROFILE), "'", "") & "' Where ID_Perfil=" & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyProfile = lErrorNumber
	Err.Clear
End Function

Function ModifyProfileAndCourses(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the asigned courses of a profile
'Inputs:  oRequest, oADODBConnection, aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyProfileAndCourses"
	Dim aCourses
	Dim aDates
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfileComponent(B_COMPONENT_INITIALIZED_PROFILE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfileComponent(oRequest, aProfileComponent)
	End If

	If aProfileComponent(N_ID_PROFILE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		aCourses = Split(aProfileComponent(S_COURSES_PROFILE), ",", -1, vbBinaryCompare)
		aStartDates = Split(aProfileComponent(S_COURSES_START_DATES_PROFILE), ",", -1, vbBinaryCompare)
		aLastDates = Split(aProfileComponent(S_COURSES_LAST_DATES_PROFILE), ",", -1, vbBinaryCompare)
		If UBound(aCourses) <> UBound(aLastDates) Then
			lErrorNumber = -1
			sErrorDescription = "El número de cursos asignados y el número de fechas límite no coinciden. La información de los cursos asignados al perfil no pudo ser modificada."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			sErrorDescription = "No se pudieron eliminar los cursos asignados al curso para modificarlos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosPerfilesLKP Where ID_Perfil=" & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				For iIndex = 0 To UBound(aCourses)
					sErrorDescription = "No se pudieron agregar los cursos al curso especificado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosPerfilesLKP (ID_Curso, ID_Perfil, Fecha_Inicio, Fecha_Final) Values (" & aCourses(iIndex) & ", " & aProfileComponent(N_ID_PROFILE) & ", " & aStartDates(iIndex) & ", " & aLastDates(iIndex) & ")", "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber <> 0 Then Exit For
				Next
				If lErrorNumber <> 0 Then
					sErrorDescription = "La información de los cursos asignados al curso no pudo ser modificada."
					If Len(Err.description) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				End If
			End If
		End If
	End If

	ModifyProfileAndCourses = lErrorNumber
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
		sErrorDescription = "No se especificó el curso a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = GetProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "Perfiles Where ID_Perfil=" & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo eliminar la información del curso."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "Perfiles Set ID_Padre=" & aProfileComponent(N_PARENT_ID_PROFILE) & " Where ID_Padre = " & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If

			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo eliminar la relación del curso eliminado con los cursos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosPerfilesLKP Where ID_Perfil = " & aProfileComponent(N_ID_PROFILE), "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
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
		sErrorDescription = "No se especificó el nombre del curso para revisar la existencia de éste en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del curso en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Perfiles Where (ID_Perfil<>" & aProfileComponent(N_ID_PROFILE) & ") And (Nombre_Perfil='" & aProfileComponent(S_NAME_PROFILE) & "')", "SADEProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aProfileComponent(B_IS_DUPLICATED_PROFILE) = True
				aProfileComponent(N_ID_PROFILE) = CLng(oRecordset.Fields("ID_Perfil").Value)
			End If
			oRecordset.Close
		End If
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
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del curso no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aProfileComponent(N_PARENT_ID_PROFILE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del curso padre no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aProfileComponent(S_NAME_PROFILE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del curso está vacío."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del curso contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADEProfileComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckProfileInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayProfilesForm(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
'************************************************************
'Purpose: To display the open profiles from SADE
'Inputs:  oRequest, oADODBConnection, aProfileComponent
'Outputs: aProfileComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfilesForm"
	Dim lErrorNumber

	If aProfileComponent(N_ID_PROFILE) <> -1 Then
		lErrorNumber = GetProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckProfileFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.ProfileName.value == '') {" & vbNewLine
							If Len(oRequest("Diploma").Item) > 0 Then
								Response.Write "alert('Favor de introducir el nombre del diplomado.');" & vbNewLine
							Else
								Response.Write "alert('Favor de introducir el nombre del curso.');" & vbNewLine
							End If
							Response.Write "oForm.ProfileName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If

				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckProfileFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<FORM NAME=""ProfileFrm"" ID=""ProfileFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckProfileFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileID"" ID=""ProfileIDHdn"" VALUE=""" & aProfileComponent(N_ID_PROFILE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileCourses"" ID=""ProfileCoursesHdn"" VALUE=""" & aProfileComponent(S_COURSES_PROFILE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileCoursesStartDates"" ID=""ProfileCoursesStartDatesHdn"" VALUE=""" & aProfileComponent(S_COURSES_START_DATES_PROFILE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileCoursesLastDates"" ID=""ProfileCoursesLastDatesHdn"" VALUE=""" & aProfileComponent(S_COURSES_LAST_DATES_PROFILE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileComments"" ID=""ProfileCommentsHdn"" VALUE=""" & aProfileComponent(S_COMMENTS_PROFILE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Diploma"" ID=""DiplomaHdn"" VALUE=""" & oRequest("Diploma").Item & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					If Len(oRequest("Diploma").Item) > 0 Then
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del diplomado:&nbsp;</NOBR></FONT></TD>"
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del curso:&nbsp;</NOBR></FONT></TD>"
					End If
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ProfileName"" ID=""ProfileNameTxt"" VALUE=""" & aProfileComponent(S_NAME_PROFILE) & """ SIZE=""30"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If Len(oRequest("Diploma").Item) > 0 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""0"" />"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Área de conocimiento:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ParentID"" ID=""ParentIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""1"">Capacitación vinculada a servicios de salud</OPTION>"
							Response.Write "<OPTION VALUE=""2"">Capacitación en apoyo a los procesos jurídicos financieros y técnico administrativo</OPTION>"
							Response.Write "<OPTION VALUE=""3"">Capacitación en tecnología de información</OPTION>"
							Response.Write "<OPTION VALUE=""4"">Capacitación pedagógica</OPTION>"
							Response.Write "<OPTION VALUE=""5"">Capacitación sobre asuntos técnico-operativo</OPTION>"
							Response.Write "<OPTION VALUE=""6"">Capacitación para la superación personal</OPTION>"
							Response.Write "<OPTION VALUE=""7"">Otros cursos</OPTION>"
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE><BR />"

			If (aProfileComponent(N_ID_PROFILE) = -1) Or (Len(oRequest("Remove").Item) > 0) Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveProfileWngDiv']); ProfileFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveProfileWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayProfilesForm = lErrorNumber
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

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileID"" ID=""ProfileIDHdn"" VALUE=""" & aProfileComponent(N_ID_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aProfileComponent(N_PARENT_ID_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileName"" ID=""ProfileNameHdn"" VALUE=""" & aProfileComponent(S_NAME_PROFILE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProfileComments"" ID=""ProfileCommentsHdn"" VALUE=""" & aProfileComponent(S_COMMENTS_PROFILE) & """ />"

	DisplayProfileAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayProfilesTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, sErrorDescription)
'************************************************************
'Purpose: To display the open profiles from SADE
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfilesTable"
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

	sErrorDescription = "No se pudo obtener la información de los cursos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & SADE_PREFIX & "Perfiles.* From " & SADE_PREFIX & "Perfiles Where (ID_Perfil>-1) " & aProfileComponent(S_QUERY_CONDITION_PROFILE) & " Order By Nombre_Perfil", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""650"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If Len(oRequest("Diploma").Item) > 0 Then
					If bUseLinks Then
						asColumnsTitles = Split("Acciones,Diplomado", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,200", ",", -1, vbBinaryCompare)
						asCellAlignments = Split("CENTER,", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("Diplomado", ",", -1, vbBinaryCompare)
						asCellWidths = Split("300", ",", -1, vbBinaryCompare)
						asCellAlignments = Split("", ",", -1, vbBinaryCompare)
					End If
				Else
					If bUseLinks Then
						asColumnsTitles = Split("Acciones,Curso,Área de conocimiento", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,200,300", ",", -1, vbBinaryCompare)
						asCellAlignments = Split("CENTER,,", ",", -1, vbBinaryCompare)
					ElseIf lIDColumn <> DISPLAY_NOTHING Then
						asColumnsTitles = Split("&nbsp;,Curso,Descripción", ",", -1, vbBinaryCompare)
						asCellWidths = Split("20,200,300", ",", -1, vbBinaryCompare)
						asCellAlignments = Split("CENTER,,", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("Curso,Descripción", ",", -1, vbBinaryCompare)
						asCellWidths = Split("200,300", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
					End If
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ID_Perfil").Value), oRequest("ProfileID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If

					sRowContents = ""
					If bUseLinks Then
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&ProfileID=" & CStr(oRecordset.Fields("ID_Perfil").Value) & "&Change=1&Diploma=" & oRequest("Diploma").Item & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&ProfileID=" & CStr(oRecordset.Fields("ID_Perfil").Value) & "&Delete=1&Diploma=" & oRequest("Diploma").Item & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR
					Else
						Select Case lIDColumn
							Case DISPLAY_RADIO_BUTTONS
								sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ProfileID"" ID=""ProfileIDRd"" VALUE=""" & CStr(oRecordset.Fields("ID_Perfil").Value) & """ />" & TABLE_SEPARATOR
							Case DISPLAY_CHECKBOXES
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ProfileID"" ID=""ProfileIDChk"" VALUE=""" & CStr(oRecordset.Fields("ID_Perfil").Value) & """ />" & TABLE_SEPARATOR
						End Select
					End If
					sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("Nombre_Perfil").Value)) & sBoldEnd
					If Len(oRequest("Diploma").Item) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin
							Select Case CInt(oRecordset.Fields("ID_Padre").Value)
								Case 1
									sRowContents = sRowContents & "Capacitación vinculada a servicios de salud"
								Case 2
									sRowContents = sRowContents & "Capacitación en apoyo a los procesos jurídicos financieros y técnico administrativo"
								Case 3
									sRowContents = sRowContents & "Capacitación en tecnología de información"
								Case 4
									sRowContents = sRowContents & "Capacitación pedagógica"
								Case 5
									sRowContents = sRowContents & "Capacitación sobre asuntos técnico-operativo"
								Case 6
									sRowContents = sRowContents & "Capacitación para la superación personal"
								Case 7
									sRowContents = sRowContents & "Otros cursos"
							End Select
						sRowContents = sRowContents & sBoldEnd
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

	Set oRecordset = Nothing
	DisplayProfilesTable = lErrorNumber
	Err.Clear
End Function
%>