<%
Const N_ID_COURSE = 0
Const S_NAME_COURSE = 1
Const S_KEY_COURSE = 2
Const S_URL_COURSE = 3
Const S_DESCRIPTION_COURSE = 4
Const N_ESTIMATED_TIME_COURSE = 5
Const N_MINUMUN_PARTICIPANTS_COURSE = 6
Const N_MAXUMUN_PARTICIPANTS_COURSE = 7
Const N_DAILY_TIME_COURSE = 8
Const D_MINIMUM_GRADE_COURSE = 9
Const N_OPTIONS_COURSE = 10
Const N_SHOW_EVALUATIONS_COURSE = 11
Const S_CERTIFICATE_COURSE = 12
Const B_ACTIVE_COURSE = 13
Const S_ID_REQUIRED_COURSE = 14
Const S_ID_GROUPS_COURSE = 15
Const S_START_DATES_COURSE = 16
Const S_LAST_DATES_COURSE = 17
Const S_ID_PROFILES_COURSE = 18

Const B_CHECK_FOR_DUPLICATED_COURSE = 21
Const B_MODIFY_COURSE = 22
Const S_QUERY_CONDITION_COURSE = 23
Const S_EVALUATIONS_COURSE = 24
Const S_EVALUATIONS_PONDERATIONS_COURSE = 25
Const S_EVALUATIONS_MINIMUM_GRADE_COURSE = 26
Const N_NUMBER_OF_EVALUATIONS_COURSE = 27
Const S_ID_COURSES_PATH_COURSE = 28
Const N_CERTIFICATE_ID_COURSE = 29
Const S_DESCRIPTOR_PATH_COURSE = 30
Const B_IS_DUPLICATED_COURSE = 31
Const B_GET_DEPENDENT_COURSES_COURSE = 32
Const S_TARGET_PAGE_COURSE = 33
Const N_ID_SELECTED_COURSE = 34
Const B_COMPONENT_INITIALIZED_COURSE = 35

Const N_COURSE_COMPONENT_SIZE = 35

Dim aCourseComponent()
ReDim aCourseComponent(N_COURSE_COMPONENT_SIZE)

Function InitializeCourseComponent(oRequest, aCourseComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Course Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aCourseComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeCourseComponent"
	Dim aGroups
	Dim aGroupsAndDates
	Dim aGroupAndDate
	Dim aProfiles
	Dim aProfilesAndDates
	Dim aProfileAndDate
	Dim iItem
	Dim iIndex
	Redim Preserve aCourseComponent(N_COURSE_COMPONENT_SIZE)

	If IsEmpty(aCourseComponent(N_ID_COURSE)) Then
		If Len(oRequest("CourseID").Item) > 0 Then
			aCourseComponent(N_ID_COURSE) = CLng(oRequest("CourseID").Item)
		Else
			aCourseComponent(N_ID_COURSE) = -1
		End If
	End If

	If IsEmpty(aCourseComponent(S_NAME_COURSE)) Then
		If Len(oRequest("CourseName").Item) > 0 Then
			aCourseComponent(S_NAME_COURSE) = oRequest("CourseName").Item
		Else
			aCourseComponent(S_NAME_COURSE) = ""
		End If
	End If
	aCourseComponent(S_NAME_COURSE) = Left(aCourseComponent(S_NAME_COURSE), 255)

	If IsEmpty(aCourseComponent(S_KEY_COURSE)) Then
		If Len(oRequest("CourseKey").Item) > 0 Then
			aCourseComponent(S_KEY_COURSE) = oRequest("CourseKey").Item
		Else
			aCourseComponent(S_KEY_COURSE) = ""
		End If
	End If
	aCourseComponent(S_KEY_COURSE) = Left(aCourseComponent(S_KEY_COURSE), 60)

	If IsEmpty(aCourseComponent(S_URL_COURSE)) Then
		If Len(oRequest("CourseURL").Item) > 0 Then
			aCourseComponent(S_URL_COURSE) = oRequest("CourseURL").Item
		Else
			aCourseComponent(S_URL_COURSE) = ""
		End If
	End If
	aCourseComponent(S_URL_COURSE) = Replace(Left(aCourseComponent(S_URL_COURSE), 255), "\", "/", 1, -1, vbBinaryCompare)

	If IsEmpty(aCourseComponent(S_DESCRIPTION_COURSE)) Then
		If Len(oRequest("CourseDescription").Item) > 0 Then
			aCourseComponent(S_DESCRIPTION_COURSE) = oRequest("CourseDescription").Item
		Else
			aCourseComponent(S_DESCRIPTION_COURSE) = ""
		End If
	End If
	aCourseComponent(S_DESCRIPTION_COURSE) = Left(aCourseComponent(S_DESCRIPTION_COURSE), 255)

	If IsEmpty(aCourseComponent(N_ESTIMATED_TIME_COURSE)) Then
		If Len(oRequest("CourseEstimatedTime").Item) > 0 Then
			aCourseComponent(N_ESTIMATED_TIME_COURSE) = CLng(oRequest("CourseEstimatedTime").Item)
		Else
			aCourseComponent(N_ESTIMATED_TIME_COURSE) = 0
		End If
	End If

	If IsEmpty(aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE)) Then
		If Len(oRequest("MinumunParticipants").Item) > 0 Then
			aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) = CLng(oRequest("MinumunParticipants").Item)
		Else
			aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) = 0
		End If
	End If

	If IsEmpty(aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE)) Then
		If Len(oRequest("MaxumunParticipants").Item) > 0 Then
			aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) = CLng(oRequest("MaxumunParticipants").Item)
		Else
			aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) = 0
		End If
	End If

	If IsEmpty(aCourseComponent(N_DAILY_TIME_COURSE)) Then
		If Len(oRequest("CourseDailyTime").Item) > 0 Then
			aCourseComponent(N_DAILY_TIME_COURSE) = CLng(oRequest("CourseDailyTime").Item)
		Else
			aCourseComponent(N_DAILY_TIME_COURSE) = -1
		End If
	End If

	If IsEmpty(aCourseComponent(D_MINIMUM_GRADE_COURSE)) Then
		If Len(oRequest("CourseMinimumGrade").Item) > 0 Then
			aCourseComponent(D_MINIMUM_GRADE_COURSE) = CDbl(oRequest("CourseMinimumGrade").Item)
		Else
			aCourseComponent(D_MINIMUM_GRADE_COURSE) = -1
		End If
	End If

	If IsEmpty(aCourseComponent(N_OPTIONS_COURSE)) Then
		If Len(oRequest("CourseOptions").Item) > 0 Then
			If InStr(1, oRequest("CourseOptions").Item, ",", vbBinaryCompare) > 1 Then
				aCourseComponent(N_OPTIONS_COURSE) = 0
				For Each iItem In oRequest("CourseOptions")
					aCourseComponent(N_OPTIONS_COURSE) = aCourseComponent(N_OPTIONS_COURSE) + CLng(iItem)
				Next
			Else
				aCourseComponent(N_OPTIONS_COURSE) = CLng(oRequest("CourseOptions").Item)
			End If
		Else
			aCourseComponent(N_OPTIONS_COURSE) = 0
		End If
	End If

	If IsEmpty(aCourseComponent(N_SHOW_EVALUATIONS_COURSE)) Then
		If Len(oRequest("ShowEvaluations").Item) > 0 Then
			aCourseComponent(N_SHOW_EVALUATIONS_COURSE) = CDbl(oRequest("ShowEvaluations").Item)
		Else
			aCourseComponent(N_SHOW_EVALUATIONS_COURSE) = 1
		End If
	End If

	If IsEmpty(aCourseComponent(S_CERTIFICATE_COURSE)) Then
		If Len(oRequest("CourseCertificate").Item) > 0 Then
			aCourseComponent(S_CERTIFICATE_COURSE) = oRequest("CourseCertificate").Item
		Else
			aCourseComponent(S_CERTIFICATE_COURSE) = ""
		End If
	End If
	aCourseComponent(S_CERTIFICATE_COURSE) = Replace(Left(aCourseComponent(S_CERTIFICATE_COURSE), 255), "\", "/", 1, -1, vbBinaryCompare)

	If IsEmpty(aCourseComponent(B_ACTIVE_COURSE)) Then
		aCourseComponent(B_ACTIVE_COURSE) = 0
		If Len(oRequest("CourseActive").Item) > 0 Then
			aCourseComponent(B_ACTIVE_COURSE) = 1
		End If
	End If

	If IsEmpty(aCourseComponent(S_ID_REQUIRED_COURSE)) Then
		If Len(oRequest("CoursesRequired").Item) > 0 Then
			aCourseComponent(S_ID_REQUIRED_COURSE) = Replace(oRequest("CoursesRequired").Item, " ", "")
		Else
			aCourseComponent(S_ID_REQUIRED_COURSE) = "-1"
		End If
	End If

	If IsEmpty(aCourseComponent(S_ID_GROUPS_COURSE)) Then
		aCourseComponent(S_ID_GROUPS_COURSE) = ""
		If Len(oRequest("GroupsAndDates").Item) > 0 Then
			aGroupsAndDates = Split(oRequest("GroupsAndDates").Item, SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
			For iIndex = 0 To UBound(aGroupsAndDates)
				aGroupAndDate = Split(aGroupsAndDates(iIndex), LIST_SEPARATOR, -1, vbBinaryCompare)
				aCourseComponent(S_ID_GROUPS_COURSE) = aCourseComponent(S_ID_GROUPS_COURSE) & aGroupAndDate(0) & ","
			Next
			If Len(aCourseComponent(S_ID_GROUPS_COURSE)) > 0 Then aCourseComponent(S_ID_GROUPS_COURSE) = Left(aCourseComponent(S_ID_GROUPS_COURSE), (Len(aCourseComponent(S_ID_GROUPS_COURSE)) - Len(",")))
		ElseIf Len(oRequest("CourseGroups").Item) > 0 Then
			aCourseComponent(S_ID_GROUPS_COURSE) = Replace(oRequest("CourseGroups").Item, " ", "")
		End If
	End If

	If IsEmpty(aCourseComponent(S_START_DATES_COURSE)) Then
		If Len(oRequest("GroupsAndDates").Item) > 0 Then
			aCourseComponent(S_START_DATES_COURSE) = ""
			For iIndex = 0 To UBound(aGroupsAndDates)
				aGroupAndDate = Split(aGroupsAndDates(iIndex), LIST_SEPARATOR, -1, vbBinaryCompare)
				aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & aGroupAndDate(1) & ","
			Next
			If Len(aCourseComponent(S_START_DATES_COURSE)) > 0 Then aCourseComponent(S_START_DATES_COURSE) = Left(aCourseComponent(S_START_DATES_COURSE), (Len(aCourseComponent(S_START_DATES_COURSE)) - Len(",")))
		ElseIf Len(oRequest("CoursesStartDate").Item) > 0 Then
			aCourseComponent(S_START_DATES_COURSE) = oRequest("CoursesStartDate").Item
		Else
			aGroups = Split(aCourseComponent(S_ID_GROUPS_COURSE), ",", -1, vbBinaryCompare)
			aCourseComponent(S_START_DATES_COURSE) = ""
			For iIndex = 0 To UBound(aGroups)
				aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & Left(GetSerialNumberForDate(""), Len("00000000")) & ","
			Next
			If Len(aCourseComponent(S_START_DATES_COURSE)) > 0 Then aCourseComponent(S_START_DATES_COURSE) = Left(aCourseComponent(S_START_DATES_COURSE), (Len(aCourseComponent(S_START_DATES_COURSE)) - Len(",")))
		End If
	End If

	If IsEmpty(aCourseComponent(S_LAST_DATES_COURSE)) Then
		If Len(oRequest("GroupsAndDates").Item) > 0 Then
			aCourseComponent(S_LAST_DATES_COURSE) = ""
			For iIndex = 0 To UBound(aGroupsAndDates)
				aGroupAndDate = Split(aGroupsAndDates(iIndex), LIST_SEPARATOR, -1, vbBinaryCompare)
				aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & aGroupAndDate(2) & ","
			Next
			If Len(aCourseComponent(S_LAST_DATES_COURSE)) > 0 Then aCourseComponent(S_LAST_DATES_COURSE) = Left(aCourseComponent(S_LAST_DATES_COURSE), (Len(aCourseComponent(S_LAST_DATES_COURSE)) - Len(",")))
		ElseIf Len(oRequest("CoursesLastDate").Item) > 0 Then
			aCourseComponent(S_LAST_DATES_COURSE) = oRequest("CoursesLastDate").Item
		Else
			aGroups = Split(aCourseComponent(S_ID_GROUPS_COURSE), ",", -1, vbBinaryCompare)
			aCourseComponent(S_LAST_DATES_COURSE) = ""
			For iIndex = 0 To UBound(aGroups)
				aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & Left(GetSerialNumberForDate(""), Len("00000000")) & ","
			Next
			If Len(aCourseComponent(S_LAST_DATES_COURSE)) > 0 Then aCourseComponent(S_LAST_DATES_COURSE) = Left(aCourseComponent(S_LAST_DATES_COURSE), (Len(aCourseComponent(S_LAST_DATES_COURSE)) - Len(",")))
		End If
	End If

	If IsEmpty(aCourseComponent(S_ID_PROFILES_COURSE)) Then
		aCourseComponent(S_ID_PROFILES_COURSE) = ""
		If Len(oRequest("ProfilesAndDates").Item) > 0 Then
			aProfilesAndDates = Split(oRequest("ProfilesAndDates").Item, SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
			For iIndex = 0 To UBound(aProfilesAndDates)
				aProfileAndDate = Split(aProfilesAndDates(iIndex), LIST_SEPARATOR, -1, vbBinaryCompare)
				aCourseComponent(S_ID_PROFILES_COURSE) = aCourseComponent(S_ID_PROFILES_COURSE) & aProfileAndDate(0) & ","
			Next
			If Len(aCourseComponent(S_ID_PROFILES_COURSE)) > 0 Then aCourseComponent(S_ID_PROFILES_COURSE) = Left(aCourseComponent(S_ID_PROFILES_COURSE), (Len(aCourseComponent(S_ID_PROFILES_COURSE)) - Len(",")))
		ElseIf Len(oRequest("CourseProfiles").Item) > 0 Then
			aCourseComponent(S_ID_PROFILES_COURSE) = Replace(oRequest("CourseProfiles").Item, " ", "")
		End If
	End If

	If IsEmpty(aCourseComponent(B_CHECK_FOR_DUPLICATED_COURSE)) Then
		aCourseComponent(B_CHECK_FOR_DUPLICATED_COURSE) = (Len(oRequest("CheckDuplicatedCourses").Item) > 0)
	End If
	aCourseComponent(B_IS_DUPLICATED_COURSE) = False

	If IsEmpty(aCourseComponent(B_MODIFY_COURSE)) Then
		aCourseComponent(B_MODIFY_COURSE) = (Len(oRequest("ModifyCourse").Item) > 0)
	End If

	If IsEmpty(aCourseComponent(S_QUERY_CONDITION_COURSE)) Then
		If Len(oRequest("CourseCondition").Item) > 0 Then
			aCourseComponent(S_QUERY_CONDITION_COURSE) = oRequest("CourseCondition").Item
		Else
			aCourseComponent(S_QUERY_CONDITION_COURSE) = ""
		End If
	End If

	If IsEmpty(aCourseComponent(S_ID_COURSES_PATH_COURSE)) Then
		If Len(oRequest("CoursesIDPath").Item) > 0 Then
			aCourseComponent(S_ID_COURSES_PATH_COURSE) = oRequest("CoursesIDPath").Item
		Else
			aCourseComponent(S_ID_COURSES_PATH_COURSE) = ""
		End If
	End If

	If IsEmpty(aCourseComponent(S_TARGET_PAGE_COURSE)) Then
		If Len(oRequest("CourseTargetPage").Item) > 0 Then
			aCourseComponent(S_TARGET_PAGE_COURSE) = oRequest("CourseTargetPage").Item
		Else
			aCourseComponent(S_TARGET_PAGE_COURSE) = GetASPFileName("")
		End If
	End If

	If IsEmpty(aCourseComponent(N_ID_SELECTED_COURSE)) Then
		If Len(oRequest("SelectedCourseID").Item) > 0 Then
			aCourseComponent(N_ID_SELECTED_COURSE) = CLng(oRequest("SelectedCourseID").Item)
		Else
			aCourseComponent(N_ID_SELECTED_COURSE) = -1
		End If
	End If

	aCourseComponent(B_COMPONENT_INITIALIZED_COURSE) = True
	InitializeCourseComponent = Err.number
	Err.Clear
End Function

Function AddCourse(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new course into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddCourse"
	Dim alRequiredCourses
	Dim alGroups
	Dim asStartDates
	Dim asLastDates
	Dim alProfiles
	Dim oItem
	Dim iIndex
	Dim sTempDate
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo curso."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, SADE_PREFIX & "Curso", "ID_Curso", "", 1, aCourseComponent(N_ID_COURSE), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aCourseComponent(B_CHECK_FOR_DUPLICATED_COURSE) Then
			lErrorNumber = CheckExistencyOfCourse(aCourseComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aCourseComponent(B_IS_DUPLICATED_COURSE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un curso registrado con el mismo nombre."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckCourseInformationConsistency(aCourseComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo guardar la información del nuevo curso."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "Curso (ID_Curso, Nombre_Curso, Clave_Curso, URL_Curso, Descripcion, Participantes_Minimo, Participantes_Maximo, TiempoEstimado, TiempoDiario, Calificacion_Minima, OpcionTerminado, MostrarEvaluaciones, Certificado, Activo) Values (" & aCourseComponent(N_ID_COURSE) & ", '" & Replace(aCourseComponent(S_NAME_COURSE), "'", "") & "', '" & Replace(aCourseComponent(S_KEY_COURSE), "'", "") & "', '" & Replace(aCourseComponent(S_URL_COURSE), "'", "") & "', '" & Replace(aCourseComponent(S_DESCRIPTION_COURSE), "'", "") & "', " & aCourseComponent(N_ESTIMATED_TIME_COURSE) & ", " & aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) & ", " & aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) & ", " & aCourseComponent(N_DAILY_TIME_COURSE) & ", " & FormatFloat(aCourseComponent(D_MINIMUM_GRADE_COURSE)) & ", " & aCourseComponent(N_OPTIONS_COURSE) & ", " & aCourseComponent(N_SHOW_EVALUATIONS_COURSE) & ", '" & Replace(aCourseComponent(S_CERTIFICATE_COURSE), "'", "") & "', " & aCourseComponent(B_ACTIVE_COURSE) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						If Len(aCourseComponent(S_ID_REQUIRED_COURSE)) > 0 Then
							alRequiredCourses = Split(aCourseComponent(S_ID_REQUIRED_COURSE), ",", -1, vbBinaryCompare)
							For Each oItem In alRequiredCourses
								If (CLng(oItem) <> -1) Or (UBound(alRequiredCourses) = 0) Then
									sErrorDescription = "No se pudo guardar la información de los cursos requeridos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursoRequisito (ID_Curso, ID_Curso_Requisito) Values (" & aCourseComponent(N_ID_COURSE) & ", " & CLng(oItem) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							Next
						Else
							sErrorDescription = "No se pudo guardar la información de los cursos requeridos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursoRequisito (ID_Curso, ID_Curso_Requisito) Values (" & aCourseComponent(N_ID_COURSE) & ", -1)", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						lErrorNumber = Err.number
						If lErrorNumber <> 0 Then
							sErrorDescription = "No se pudieron asociar los cursos requisitos con la información del nuevo curso."
							If Len(Err.description) > 0 Then
								sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
							End If
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
						End If

						If Len(aCourseComponent(S_ID_GROUPS_COURSE)) > 0 Then
							alGroups = Split(aCourseComponent(S_ID_GROUPS_COURSE), ",", -1, vbBinaryCompare)
							asStartDates = Split(aCourseComponent(S_START_DATES_COURSE), ",", -1, vbBinaryCompare)
							asLastDates = Split(aCourseComponent(S_LAST_DATES_COURSE), ",", -1, vbBinaryCompare)
							aCourseComponent(S_START_DATES_COURSE) = ""
							aCourseComponent(S_LAST_DATES_COURSE) = ""
							For iIndex = 0 To UBound(alGroups)
								If CLng(alGroups(iIndex)) <> -1 Then
									sErrorDescription = "No se pudo guardar la información de los grupos que tomarán este curso."
									If iIndex > UBound(asLastDates) Then
										sTempDate = Left(GetSerialNumberForDate(""), Len("00000000"))
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosGruposLKP (ID_Curso, ID_Grupo, Fecha_Inicio, Fecha_Final) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alGroups(iIndex) & ", " & sTempDate & ", " & sTempDate & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If (lErrorNumber = 0) And B_HISTORY_LIST Then lErrorNumber = AddRegisterIntoHistoryList(oADODBConnection, -1, alGroups(iIndex), -1, aCourseComponent(N_ID_COURSE), sTempDate, sTempDate, 0, sErrorDescription)
										aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & sTempDate & ","
										aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & sTempDate & ","
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosGruposLKP (ID_Curso, ID_Grupo, Fecha_Inicio, Fecha_Final) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alGroups(iIndex) & ", " & asStartDates(iIndex) & ", " & asLastDates(iIndex) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If (lErrorNumber = 0) And B_HISTORY_LIST Then lErrorNumber = AddRegisterIntoHistoryList(oADODBConnection, -1, alGroups(iIndex), -1, aCourseComponent(N_ID_COURSE), asStartDates(iIndex), asLastDates(iIndex), 0, sErrorDescription)
										aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & asStartDates(iIndex) & ","
										aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & asLastDates(iIndex) & ","
									End If
								End If
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							Next
							If Len(aCourseComponent(S_START_DATES_COURSE)) > 0 Then
								aCourseComponent(S_START_DATES_COURSE) = Left(aCourseComponent(S_START_DATES_COURSE), (Len(aCourseComponent(S_START_DATES_COURSE)) - Len(",")))
							End If
							If Len(aCourseComponent(S_LAST_DATES_COURSE)) > 0 Then
								aCourseComponent(S_LAST_DATES_COURSE) = Left(aCourseComponent(S_LAST_DATES_COURSE), (Len(aCourseComponent(S_LAST_DATES_COURSE)) - Len(",")))
							End If
						End If

						If Len(aCourseComponent(S_ID_PROFILES_COURSE)) > 0 Then
							alProfiles = Split(aCourseComponent(S_ID_PROFILES_COURSE), ",", -1, vbBinaryCompare)
							For iIndex = 0 To UBound(alProfiles)
								If CLng(alProfiles(iIndex)) <> -1 Then
									sErrorDescription = "No se pudo guardar la información de los perfiles que tomarán este curso."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosPerfilesLKP (ID_Curso, ID_Perfil) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alProfiles(iIndex) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							Next
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddCourse = lErrorNumber
	Err.Clear
End Function

Function GetCourse(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a course from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCourse"
	Dim sFolder
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del curso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Curso Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El curso especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aCourseComponent(S_NAME_COURSE) = CStr(oRecordset.Fields("Nombre_Curso").Value)
				aCourseComponent(S_KEY_COURSE) = CStr(oRecordset.Fields("Clave_Curso").Value)
				aCourseComponent(S_URL_COURSE) = CStr(oRecordset.Fields("URL_Curso").Value)
				aCourseComponent(S_DESCRIPTION_COURSE) = CStr(oRecordset.Fields("Descripcion").Value)
				aCourseComponent(N_ESTIMATED_TIME_COURSE) = CLng(oRecordset.Fields("TiempoEstimado").Value)
				aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) = CLng(oRecordset.Fields("Participantes_Minimo").Value)
				aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) = CLng(oRecordset.Fields("Participantes_Maximo").Value)
				aCourseComponent(N_DAILY_TIME_COURSE) = CLng(oRecordset.Fields("TiempoDiario").Value)
				aCourseComponent(D_MINIMUM_GRADE_COURSE) = CDbl(oRecordset.Fields("Calificacion_Minima").Value)
				aCourseComponent(N_OPTIONS_COURSE) = CLng(oRecordset.Fields("OpcionTerminado").Value)
				aCourseComponent(N_SHOW_EVALUATIONS_COURSE) = CInt(oRecordset.Fields("MostrarEvaluaciones").Value)
				aCourseComponent(S_CERTIFICATE_COURSE) = CStr(oRecordset.Fields("Certificado").Value)
				aCourseComponent(B_ACTIVE_COURSE) = CInt(oRecordset.Fields("Activo").Value)
				sFolder = Replace((COURSES_PATH & aCourseComponent(S_URL_COURSE)), "/", "\", 1, -1, vbBinaryCompare)
				sFolder = Left(sFolder, InStrRev(sFolder, "\")) & "Descriptor.xml"
				If InStr(1, Server.MapPath(".\Descriptor.xml"), sFolder, vbBinaryCompare) = 0 Then
					aCourseComponent(S_DESCRIPTOR_PATH_COURSE) = Server.MapPath(sFolder)
				Else
					aCourseComponent(S_DESCRIPTOR_PATH_COURSE) = Server.MapPath(".\Descriptor.xml")
				End If
			End If
			oRecordset.Close

			If lErrorNumber = 0 Then
				If (Len(aCourseComponent(S_ID_REQUIRED_COURSE)) = 0) Or (StrComp(aCourseComponent(S_ID_REQUIRED_COURSE), "-1", vbBinaryCompare) = 0) Then
					sErrorDescription = "No se pudieron obtener los cursos requisitos para el curso especificado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ID_Curso_Requisito From " & SADE_PREFIX & "CursoRequisito Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						aCourseComponent(S_ID_REQUIRED_COURSE) = "-1,"
						If Not oRecordset.EOF Then
							Do While Not oRecordset.EOF
								If CLng(oRecordset.Fields("ID_Curso_Requisito").Value) <> -1 Then aCourseComponent(S_ID_REQUIRED_COURSE) = aCourseComponent(S_ID_REQUIRED_COURSE) & CStr(oRecordset.Fields("ID_Curso_Requisito").Value) & ","
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
						oRecordset.Close
						aCourseComponent(S_ID_REQUIRED_COURSE) = Left(aCourseComponent(S_ID_REQUIRED_COURSE), Len(aCourseComponent(S_ID_REQUIRED_COURSE)) - Len(","))
					End If
				End If

				If Len(aCourseComponent(S_ID_GROUPS_COURSE)) = 0 Then
					sErrorDescription = "No se pudieron obtener los grupos para el curso especificado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ID_Grupo, Fecha_Inicio, Fecha_Final From " & SADE_PREFIX & "CursosGruposLKP Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							aCourseComponent(S_ID_GROUPS_COURSE) = "-1,"
							aCourseComponent(S_START_DATES_COURSE) = Left(GetSerialNumberForDate(""), Len("00000000")) & ","
							aCourseComponent(S_LAST_DATES_COURSE) = Left(GetSerialNumberForDate(""), Len("00000000")) & ","
						Else
							aCourseComponent(S_ID_GROUPS_COURSE) = ""
							aCourseComponent(S_START_DATES_COURSE) = ""
							aCourseComponent(S_LAST_DATES_COURSE) = ""
							Do While Not oRecordset.EOF
								aCourseComponent(S_ID_GROUPS_COURSE) = aCourseComponent(S_ID_GROUPS_COURSE) & CStr(oRecordset.Fields("ID_Grupo").Value) & ","
								aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & CStr(oRecordset.Fields("Fecha_Inicio").Value) & ","
								aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & CStr(oRecordset.Fields("Fecha_Final").Value) & ","
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
						oRecordset.Close
						aCourseComponent(S_ID_GROUPS_COURSE) = Left(aCourseComponent(S_ID_GROUPS_COURSE), Len(aCourseComponent(S_ID_GROUPS_COURSE)) - Len(","))
						aCourseComponent(S_START_DATES_COURSE) = Left(aCourseComponent(S_START_DATES_COURSE), Len(aCourseComponent(S_START_DATES_COURSE)) - Len(","))
						aCourseComponent(S_LAST_DATES_COURSE) = Left(aCourseComponent(S_LAST_DATES_COURSE), Len(aCourseComponent(S_LAST_DATES_COURSE)) - Len(","))
					End If
				End If

				If Len(aCourseComponent(S_ID_PROFILES_COURSE)) = 0 Then
					sErrorDescription = "No se pudieron obtener los perfiles para el curso especificado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ID_Perfil From " & SADE_PREFIX & "CursosPerfilesLKP Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							aCourseComponent(S_ID_PROFILES_COURSE) = "-1,"
						Else
							aCourseComponent(S_ID_PROFILES_COURSE) = ""
							Do While Not oRecordset.EOF
								aCourseComponent(S_ID_PROFILES_COURSE) = aCourseComponent(S_ID_PROFILES_COURSE) & CStr(oRecordset.Fields("ID_Perfil").Value) & ","
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
						oRecordset.Close
						aCourseComponent(S_ID_PROFILES_COURSE) = Left(aCourseComponent(S_ID_PROFILES_COURSE), Len(aCourseComponent(S_ID_PROFILES_COURSE)) - Len(","))
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetCourse = lErrorNumber
	Err.Clear
End Function

Function GetCourses(oRequest, oADODBConnection, aCourseComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the courses from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCourses"
	Dim sCondition
	Dim sColumns
	Dim sTables
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	sCondition = ""
	If Len(aCourseComponent(S_QUERY_CONDITION_COURSE)) > 0 Then
		aCourseComponent(S_QUERY_CONDITION_COURSE) = LTrim(aCourseComponent(S_QUERY_CONDITION_COURSE))
		If InStr(1, aCourseComponent(S_QUERY_CONDITION_COURSE), "And", vbTextCompare) = 1 Then
			aCourseComponent(S_QUERY_CONDITION_COURSE) = Replace(aCourseComponent(S_QUERY_CONDITION_COURSE), "And", "", 1, 1, vbTextCompare)
		End If
		sCondition = " Where " & aCourseComponent(S_QUERY_CONDITION_COURSE)
	End If
	sColumns = SADE_PREFIX & "Curso.*"
	sTables = SADE_PREFIX & "Curso"
	If InStr(1, sCondition, "CursosGruposLKP", vbTextCompare) > 0 Then
		sTables = sTables & ", " & SADE_PREFIX & "CursosGruposLKP"
		sColumns = sColumns & ", " & SADE_PREFIX & "CursosGruposLKP.*"
	End If
	If InStr(1, sCondition, "CursosPerfilesLKP", vbTextCompare) > 0 Then
		sTables = sTables & ", " & SADE_PREFIX & "CursosPerfilesLKP"
		sColumns = sColumns & ", " & SADE_PREFIX & "CursosPerfilesLKP.*"
	End If
	sErrorDescription = "No se pudo obtener la información de los cursos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sColumns & " From " & sTables & " " & sCondition & " Order By Nombre_Curso", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetCourses = lErrorNumber
	Err.Clear
End Function

Function ModifyCourse(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing course in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyCourse"
	Dim alRequiredCourses
	Dim alGroups
	Dim asStartDates
	Dim asLastDates
	Dim alProfiles
	Dim iIndex
	Dim sTempDate
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckCourseInformationConsistency(aCourseComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo actualizar la información del curso especificado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "Curso Set Nombre_Curso='" & Replace(aCourseComponent(S_NAME_COURSE), "'", "") & "', Clave_Curso='" & Replace(aCourseComponent(S_KEY_COURSE), "'", "") & "', URL_Curso='" & Replace(aCourseComponent(S_URL_COURSE), "'", "") & "', Descripcion='" & Replace(aCourseComponent(S_DESCRIPTION_COURSE), "'", "") & "', TiempoEstimado=" & aCourseComponent(N_ESTIMATED_TIME_COURSE) & ", Participantes_Minimo=" & aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) & ", Participantes_Maximo=" & aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) & ", TiempoDiario=" & aCourseComponent(N_DAILY_TIME_COURSE) & ", Calificacion_Minima=" & FormatFloat(aCourseComponent(D_MINIMUM_GRADE_COURSE)) & ", OpcionTerminado=" & aCourseComponent(N_OPTIONS_COURSE) & ", MostrarEvaluaciones=" & aCourseComponent(N_SHOW_EVALUATIONS_COURSE) & ", Certificado='" & Replace(aCourseComponent(S_CERTIFICATE_COURSE), "'", "") & "', Activo=" & aCourseComponent(B_ACTIVE_COURSE) & " Where ID_Curso = " & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Call RemoveRequiredCourses(oRequest, oADODBConnection, aCourseComponent, False, "")
			If Len(aCourseComponent(S_ID_REQUIRED_COURSE)) > 0 Then
				alRequiredCourses = Split(aCourseComponent(S_ID_REQUIRED_COURSE), ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(alRequiredCourses)
					If CLng(alRequiredCourses(iIndex)) <> -1 Then
						sErrorDescription = "No se pudieron relacionar los cursos requeridos con el curso especificado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursoRequisito (ID_Curso, ID_Curso_Requisito) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alRequiredCourses(iIndex) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If (Err.number <> 0) Or (lErrorNumber) Then Exit For
				Next
			End If
			If (StrComp(aCourseComponent(S_ID_REQUIRED_COURSE), "-1", vbBinaryCompare) = 0) Or (Len(aCourseComponent(S_ID_REQUIRED_COURSE)) = 0) Then
				sErrorDescription = "No se pudieron indicar que el curso especificado no tiene cursos requeridos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursoRequisito (ID_Curso, ID_Curso_Requisito) Values (" & aCourseComponent(N_ID_COURSE) & ", -1)", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If

			Call RemoveGroups(oRequest, oADODBConnection, aCourseComponent, "")
			If Len(aCourseComponent(S_ID_GROUPS_COURSE)) > 0 Then
				alGroups = Split(aCourseComponent(S_ID_GROUPS_COURSE), ",", -1, vbBinaryCompare)
				asStartDates = Split(aCourseComponent(S_START_DATES_COURSE), ",", -1, vbBinaryCompare)
				asLastDates = Split(aCourseComponent(S_LAST_DATES_COURSE), ",", -1, vbBinaryCompare)
				aCourseComponent(S_START_DATES_COURSE) = ""
				aCourseComponent(S_LAST_DATES_COURSE) = ""
				For iIndex = 0 To UBound(alGroups)
					If CLng(alGroups(iIndex)) <> -1 Then
						sErrorDescription = "No se pudo guardar la información de los grupos que tomarán este curso."
						If iIndex > UBound(asLastDates) Then
							sTempDate = Left(GetSerialNumberForDate(""), Len("00000000"))
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosGruposLKP (ID_Curso, ID_Grupo, Fecha_Inicio, Fecha_Final) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alGroups(iIndex) & ", " & sTempDate & ", " & sTempDate & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & sTempDate  & ","
							aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & sTempDate  & ","
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosGruposLKP (ID_Curso, ID_Grupo, Fecha_Inicio, Fecha_Final) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alGroups(iIndex) & ", " & asStartDates(iIndex) & ", " & asLastDates(iIndex) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							aCourseComponent(S_START_DATES_COURSE) = aCourseComponent(S_START_DATES_COURSE) & asStartDates(iIndex) & ","
							aCourseComponent(S_LAST_DATES_COURSE) = aCourseComponent(S_LAST_DATES_COURSE) & asLastDates(iIndex) & ","
						End If
					End If
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
				Next
				If Len(aCourseComponent(S_START_DATES_COURSE)) > 0 Then
					aCourseComponent(S_START_DATES_COURSE) = Left(aCourseComponent(S_START_DATES_COURSE), (Len(aCourseComponent(S_START_DATES_COURSE)) - Len(",")))
				End If
				If Len(aCourseComponent(S_LAST_DATES_COURSE)) > 0 Then
					aCourseComponent(S_LAST_DATES_COURSE) = Left(aCourseComponent(S_LAST_DATES_COURSE), (Len(aCourseComponent(S_LAST_DATES_COURSE)) - Len(",")))
				End If
			End If

			Call RemoveProfiles(oRequest, oADODBConnection, aCourseComponent, "")
			If Len(aCourseComponent(S_ID_PROFILES_COURSE)) > 0 Then
				alProfiles = Split(aCourseComponent(S_ID_PROFILES_COURSE), ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(alProfiles)
					If CLng(alProfiles(iIndex)) <> -1 Then
						sErrorDescription = "No se pudo guardar la información de los perfiles que tomarán este curso."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosPerfilesLKP (ID_Curso, ID_Perfil) Values (" & aCourseComponent(N_ID_COURSE) & ", " & alProfiles(iIndex) & ")", "SADECourseComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
				Next
			End If

			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudo modificar la información del curso."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	End If

	ModifyCourse = lErrorNumber
	Err.Clear
End Function

Function RemoveCourse(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a course from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveCourse"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el curso a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del curso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "Curso Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron eliminar las entradas de los usuarios al curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "EntradasCursos Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)

			Call RemoveRequiredCourses(oRequest, oADODBConnection, aCourseComponent, True, "")
			Call RemoveGroups(oRequest, oADODBConnection, aCourseComponent, "")
			Call RemoveProfiles(oRequest, oADODBConnection, aCourseComponent, "")
		End If
	End If

	RemoveCourse = lErrorNumber
	Err.Clear
End Function

Function RemoveRequiredCourses(oRequest, oADODBConnection, aCourseComponent, bRemovingCourse, sErrorDescription)
'************************************************************
'Purpose: To remove the required courses asociated with a course
'Inputs:  oRequest, oADODBConnection, bRemovingCourse
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveRequiredCourses"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso para eliminar sus cursos requeridos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudieron eliminar los cursos requeridos para el curso especificado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursoRequisito Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
		If bRemovingCourse Then
			sErrorDescription = "No se pudo eliminar la relación de los cursos dependientes con el curso especificado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "CursoRequisito Set ID_Curso_Requisito=-1 Where ID_Curso_Requisito=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
		End If
	End If

	RemoveRequiredCourses = lErrorNumber
	Err.Clear
End Function

Function RemoveGroups(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To remove the asociation between groups and the course
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveGroups"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso para eliminar su asociación con los grupos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la asociación de los grupos con el curso especificado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosGruposLKP Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
	End If

	RemoveGroups = lErrorNumber
	Err.Clear
End Function

Function RemoveProfiles(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To remove the asociation between profiles and the course
'Inputs:  oRequest, oADODBConnection
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveProfiles"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If aCourseComponent(N_ID_COURSE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del curso para eliminar su asociación con los perfiles."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la asociación de los perfiles con el curso especificado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosPerfilesLKP Where ID_Curso=" & aCourseComponent(N_ID_COURSE), "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
	End If

	RemoveProfiles = lErrorNumber
	Err.Clear
End Function

Function SaveCertificate(oRequest, oADODBConnection, lUserID, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To save a new registry in the certificate table
'Inputs:  oRequest, oADODBConnection, lUserID, aCourseComponent
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SaveCertificate"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	sErrorDescription = "No se pudo revisar la existencia del certificado impreso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ID_Constancia From " & SADE_PREFIX & "Constancias Where (ID_Usuario=" & lUserID & ") And (ID_Curso=" & aCourseComponent(N_ID_COURSE) & ")", "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		aCourseComponent(N_CERTIFICATE_ID_COURSE) = CLng(oRecordset.Fields("ID_Constancia").Value)
	Else
		sErrorDescription = "No se pudo obtener un identificador para la constancia."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, SADE_PREFIX & "Constancias", "ID_Constancia", "", 1, aCourseComponent(N_CERTIFICATE_ID_COURSE), sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo registrar la impresión del certificado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "Constancias (ID_Constancia, ID_Usuario, ID_Curso) Values (" & aCourseComponent(N_CERTIFICATE_ID_COURSE) & ", " & lUserID & ", " & aCourseComponent(N_ID_COURSE) & ")", "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, Null)
		End If
	End If

	SaveCertificate = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfCourse(aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific course exists in the database
'Inputs:  aCourseComponent
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfCourse"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCourseComponent(B_COMPONENT_INITIALIZED_COURSE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCourseComponent(oRequest, aCourseComponent)
	End If

	If Len(aCourseComponent(S_NAME_COURSE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del curso para revisar la existencia de éste en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del curso en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Curso Where Nombre_Curso='" & aCourseComponent(S_NAME_COURSE) & "'", "SADECourseComponent.asp", 000, S_FUNCTION_NAME, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aCourseComponent(B_IS_DUPLICATED_COURSE) = True
				aCourseComponent(N_ID_COURSE) = CLng(oRecordset.Fields("ID_Curso").Value)
				aCourseComponent(B_ACTIVE_COURSE) = CInt(oRecordset.Fields("Activo").Value)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfCourse = lErrorNumber
	Err.Clear
End Function

Function CheckCourseInformationConsistency(aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aCourseComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckCourseInformationConsistency"
	Dim bIsCorrect
	Dim alGroups
	Dim asDates
	Dim alProfiles
	Dim asDatesProfiles
	Dim alRequiredCourses
	Dim oItem

	bIsCorrect = True

	If Not IsNumeric(aCourseComponent(N_ID_COURSE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del curso no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aCourseComponent(S_NAME_COURSE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del curso está vacío."
		bIsCorrect = False
	End If
	If Len(aCourseComponent(S_KEY_COURSE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave del curso está vacía."
		bIsCorrect = False
	End If
	If Len(aCourseComponent(S_URL_COURSE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del archivo donde inicia el curso está vacío."
		bIsCorrect = False
	End If
	alGroups = Split(aCourseComponent(S_ID_GROUPS_COURSE), ",", -1, vbBinaryCompare)
	For Each oItem In alGroups
		If Not IsNumeric(oItem) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- Uno de los identificadores de los grupos para el curso no es un valor numérico."
			bIsCorrect = False
			Exit For
		End If
	Next
	asDates = Split(aCourseComponent(S_START_DATES_COURSE), ",", -1, vbBinaryCompare)
	For Each oItem In asDates
		If Not IsNumeric(oItem) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- Una de las fechas de inicio del curso no es válida."
			bIsCorrect = False
			Exit For
		End If
	Next
	asDates = Split(aCourseComponent(S_LAST_DATES_COURSE), ",", -1, vbBinaryCompare)
	For Each oItem In asDates
		If Not IsNumeric(oItem) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- Una de las fechas para completar el curso no es válida."
			bIsCorrect = False
			Exit For
		End If
	Next
	alProfiles = Split(aCourseComponent(S_ID_PROFILES_COURSE), ",", -1, vbBinaryCompare)
	For Each oItem In alProfiles
		If Not IsNumeric(oItem) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- Uno de los identificadores de los perfiles para el curso no es un valor numérico."
			bIsCorrect = False
			Exit For
		End If
	Next
	alRequiredCourses = Split(aCourseComponent(S_ID_REQUIRED_COURSE), ",", -1, vbBinaryCompare)
	For Each oItem In alRequiredCourses
		If Not IsNumeric(oItem) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- Uno de los identificadores de los cursos requisito no es un valor numérico."
			bIsCorrect = False
			Exit For
		End If
	Next

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del curso contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "SADECourseComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckCourseInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayCoursesForm(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To display the open courses from SADE
'Inputs:  oRequest, oADODBConnection, aCourseComponent
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCoursesForm"
	Dim lErrorNumber

	If aCourseComponent(N_ID_COURSE) <> -1 Then
		lErrorNumber = GetCourse(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckCourseFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.CourseName.value == '') {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre del curso.');" & vbNewLine
							Response.Write "oForm.CourseName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (!CheckIntegerValue(oForm.CourseEstimatedTime, 'la duración del curso', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oForm.CoursesStartDate.value = oForm.CoursesStartYear.value + oForm.CoursesStartMonth.value + oForm.CoursesStartDay.value;" & vbNewLine
						Response.Write "oForm.CoursesLastDate.value = oForm.CoursesLastYear.value + oForm.CoursesLastMonth.value + oForm.CoursesLastDay.value;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If

				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckCourseFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<FORM NAME=""CourseFrm"" ID=""CourseFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckCourseFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & aCourseComponent(N_ID_COURSE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseKey"" ID=""CourseKeyHdn"" VALUE=""."" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseURL"" ID=""CourseURLHdn"" VALUE=""index.htm"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseDailyTime"" ID=""CourseDailyTimeHdn"" VALUE=""-1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseMinimumGrade"" ID=""CourseMinimumGradeHdn"" VALUE=""-1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseOptions"" ID=""CourseOptionsHdn"" VALUE=""0"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseCertificate"" ID=""CourseCertificateHdn"" VALUE=""" & aCourseComponent(S_CERTIFICATE_COURSE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseActive"" ID=""CourseActiveHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CoursesRequired"" ID=""CoursesRequiredHdn"" VALUE=""-1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseGroups"" ID=""CourseGroupsHdn"" VALUE=""-2"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CoursesStartDate"" ID=""CoursesStartDateHdn"" VALUE=""0"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CoursesLastDate"" ID=""CoursesLastDateHdn"" VALUE=""0"" />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del curso:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CourseName"" ID=""CourseNameTxt"" VALUE=""" & aCourseComponent(S_NAME_COURSE) & """ SIZE=""30"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Diplomado:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ShowEvaluations"" ID=""ShowEvaluationsCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "Perfiles", "ID_Perfil", "Nombre_Perfil", "(ID_Padre=0)", "Nombre_Perfil", aCourseComponent(N_SHOW_EVALUATIONS_COURSE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
 				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Lugar donde se impartirá:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CourseDescription"" ID=""CourseDescriptionTxt"" VALUE=""" & aCourseComponent(S_DESCRIPTION_COURSE) & """ SIZE=""30"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Duración:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CourseEstimatedTime"" ID=""CourseEstimatedTimeTxt"" VALUE=""" & aCourseComponent(N_ESTIMATED_TIME_COURSE) & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">&nbsp;horas</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>No. mínimo de participantes:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MinumunParticipants"" ID=""MinumunParticipantsTxt"" VALUE=""" & aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) & """ SIZE=""3"" MAXLENGTH=""3"" CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">&nbsp;personas</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>No. máximo de participantes:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MaxumunParticipants"" ID=""MaxumunParticipantsTxt"" VALUE=""" & aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) & """ SIZE=""3"" MAXLENGTH=""3"" CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">&nbsp;personas</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de inicio:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aCourseComponent(S_START_DATES_COURSE), "CoursesStart", Year(Date()), Year(Date()) + 2, True, False)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de término:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aCourseComponent(S_LAST_DATES_COURSE), "CoursesLast", Year(Date()), Year(Date()) + 2, True, False)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If Len(aCourseComponent(S_CERTIFICATE_COURSE)) > 0 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><A HREF=""" & UPLOADED_PATH & "Courses/" & aCourseComponent(S_CERTIFICATE_COURSE) & """ TARGET=""Courses""><B>Ficha técnica del curso</B></A><BR /><BR /></FONT>"
			End If
			Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=Courses&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME><BR />"

			If (aCourseComponent(N_ID_COURSE) = -1) Or (Len(oRequest("Remove").Item) > 0) Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveCourseWngDiv']); CourseFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveCourseWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayCoursesForm = lErrorNumber
	Err.Clear
End Function

Function DisplayCourseAsHiddenFields(oRequest, oADODBConnection, aCourseComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a course using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aCourseComponent
'Outputs: aCourseComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCourseAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & aCourseComponent(N_ID_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseName"" ID=""CourseNameHdn"" VALUE=""" & aCourseComponent(S_NAME_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseKey"" ID=""CourseKeyHdn"" VALUE=""" & aCourseComponent(S_KEY_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseURL"" ID=""CourseURLHdn"" VALUE=""" & aCourseComponent(S_URL_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseDescription"" ID=""CourseDescriptionHdn"" VALUE=""" & aCourseComponent(S_DESCRIPTION_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseEstimatedTime"" ID=""CourseEstimatedTimeHdn"" VALUE=""" & aCourseComponent(N_ESTIMATED_TIME_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MinumunParticipants"" ID=""MinumunParticipantsHdn"" VALUE=""" & aCourseComponent(N_MINUMUN_PARTICIPANTS_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MaxumunParticipants"" ID=""MaxumunParticipantsHdn"" VALUE=""" & aCourseComponent(N_MAXUMUN_PARTICIPANTS_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseDailyTime"" ID=""CourseDailyTimeHdn"" VALUE=""" & aCourseComponent(N_DAILY_TIME_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseMinimumGrade"" ID=""CourseMinimumGradeHdn"" VALUE=""" & aCourseComponent(D_MINIMUM_GRADE_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseOptions"" ID=""CourseOptionsHdn"" VALUE=""" & aCourseComponent(N_OPTIONS_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowEvaluations"" ID=""ShowEvaluationsHdn"" VALUE=""" & aCourseComponent(N_SHOW_EVALUATIONS_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseCertificate"" ID=""CourseCertificateHdn"" VALUE=""" & aCourseComponent(S_CERTIFICATE_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseActive"" ID=""CourseActiveHdn"" VALUE=""" & aCourseComponent(B_ACTIVE_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CoursesRequired"" ID=""CoursesRequiredHdn"" VALUE=""" & aCourseComponent(S_ID_REQUIRED_COURSE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupsAndDates"" ID=""GroupsAndDatesHdn"" VALUE=""" & aCourseComponent(S_ID_GROUPS_COURSE) & LIST_SEPARATOR & aCourseComponent(S_START_DATES_COURSE) & LIST_SEPARATOR & aCourseComponent(S_LAST_DATES_COURSE) & """ />"

	DisplayCourseAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayCoursesTable(oRequest, oADODBConnection, bAll, lIDColumn, bUseLinks, sErrorDescription)
'************************************************************
'Purpose: To display the open courses from SADE
'Inputs:  oRequest, oADODBConnection, bAll, lIDColumn, bUseLinks
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCoursesTable"
	Dim sCondition
	Dim oRecordset
	Dim sTemp
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber

	sCondition = ""
	If Not bAll Then sCondition = "And (Fecha_Final>=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
	sErrorDescription = "No se pudo obtener la información de los cursos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & SADE_PREFIX & "Curso.*, ID_Perfil, Nombre_Perfil, Fecha_Inicio, Fecha_Final From " & SADE_PREFIX & "Curso, " & SADE_PREFIX & "Perfiles, " & SADE_PREFIX & "CursosGruposLKP Where (" & SADE_PREFIX & "Curso.MostrarEvaluaciones=" & SADE_PREFIX & "Perfiles.ID_Perfil) And (" & SADE_PREFIX & "Curso.ID_Curso=" & SADE_PREFIX & "CursosGruposLKP.ID_Curso) " & sCondition & " Order By Nombre_Perfil, Nombre_Curso", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""650"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("Acciones,Diplomado,Curso,Fecha de inicio,Fecha de término,Lugar", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,100,250", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
				ElseIf lIDColumn <> DISPLAY_NOTHING Then
					asColumnsTitles = Split("&nbsp;,Diplomado,Curso,Fecha de inicio,Fecha de término,Lugar", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,100,100,100,250", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Diplomado,Curso,Fecha de inicio,Fecha de término,Lugar", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,250", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ID_Curso").Value), oRequest("CourseID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If

					sRowContents = ""
					If bUseLinks Then
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&CourseID=" & CStr(oRecordset.Fields("ID_Curso").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&CourseID=" & CStr(oRecordset.Fields("ID_Curso").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR
					Else
						Select Case lIDColumn
							Case DISPLAY_RADIO_BUTTONS
								sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""CourseID"" ID=""CourseIDRd"" VALUE=""" & CStr(oRecordset.Fields("ID_Curso").Value) & """ />" & TABLE_SEPARATOR
							Case DISPLAY_CHECKBOXES
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""CourseID"" ID=""CourseIDChk"" VALUE=""" & CStr(oRecordset.Fields("ID_Curso").Value) & """ />" & TABLE_SEPARATOR
						End Select
					End If
					If (CLng(oRecordset.Fields("ID_Perfil").Value)) > -1 Then
						sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("Nombre_Perfil").Value))
					Else
						sRowContents = sRowContents & sBoldBegin & "<CENTER>---</CENTER>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "<A"
						sTemp = ""
						sTemp = CStr(oRecordset.Fields("Certificado").Value)
						Err.Clear
						If Len(sTemp) > 0 Then sRowContents = sRowContents & " HREF=""" & UPLOADED_PATH & "Courses/" & sTemp & """ TARGET=""Courses"""
					sRowContents = sRowContents & ">" & CleanStringForHTML(CStr(oRecordset.Fields("Nombre_Curso").Value)) & "</A>" & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Inicio").Value), -1, -1, -1) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Final").Value), -1, -1, -1) & sBoldEnd
					sTemp = ""
					sTemp = CStr(oRecordset.Fields("Descripcion").Value)
					Err.Clear
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(sTemp) & sBoldEnd

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
	DisplayCoursesTable = lErrorNumber
	Err.Clear
End Function
%>