<!-- #include file="UploadInfoDisplayLibrary.asp" -->
<%
Function SaveEmployeeChildren(aEmployeeComponent, sAction, sErrorDescription)
'************************************************************
'Purpose: To check the child information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent, sAction
'Outputs: bExistChild
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SaveEmployeeChildren"
	Dim bExistChild
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo almacenar la información de los hijos de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ChildID, LevelID From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildName= '" & Replace(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE), "'", "´") & "') And (ChildLastName= '" & Replace(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE), "'", "´") & "') And (ChildBirthDate=" & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			If StrComp(sAction, "EmployeesChildren", vbBinaryCompare) = 0 Then
				lErrorNumber = GetEmployeeChildren(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If Len(aEmployeeComponent(S_NAME_EMPLOYEE)) > 0 Then
					aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) = -1
					lErrorNumber = AddEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Else
					lErrorNumber = -1
				End If
			Else
				lErrorNumber = -1
			End If
		Else
			aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildID").Value)
			If StrComp(sAction, "EmployeesChildren", vbBinaryCompare) = 0 Then
				aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) = CInt(oRecordset.Fields("ChildID").Value)
			End If
			lErrorNumber = ModifyEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	SaveEmployeeChildren = lErrorNumber
	Err.Clear
End Function

Function ShowUploadedFile(sFileName, iColumns, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file and its columns
'Inputs:  sFileName
'Outputs: iColumns, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowUploadedFile"
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim iIndex
	Dim jIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "El archivo está vacío."
	Else
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
			For jIndex = 0 To UBound(asFileRow)
				sRowContents = sRowContents & "Columna " & (jIndex + 1)
				If jIndex < UBound(asFileRow) Then sRowContents = sRowContents & TABLE_SEPARATOR
			Next
			iColumns = jIndex
			asColumnsTitles = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					sRowContents = ""
					For jIndex = 0 To UBound(asFileRow)
						sRowContents = sRowContents & CleanStringForHTML(asFileRow(jIndex))
						If jIndex < UBound(asFileRow) Then sRowContents = sRowContents & TABLE_SEPARATOR
					Next
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				If iIndex >= 9 Then Exit For
			Next
		Response.Write "</TABLE><BR />"
	End If

	ShowUploadedFile = lErrorNumber
	Err.Clear
End Function

Function UploadChildrenSchoolarshipsFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesChildrenLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadChildrenSchoolarshipsFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "ChildName"
						aiFieldsOrder(1) = iIndex & ",ChildName"
					Case "ChildLastName"
						aiFieldsOrder(2) = iIndex & ",ChildLastName"
					Case "ChildLastName2"
						aiFieldsOrder(3) = iIndex & ",ChildLastName2"
					Case "LevelID"
						aiFieldsOrder(4) = iIndex & ",LevelID"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(5) = iIndex & ",ChildBirthDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(5) = iIndex & ",ChildBirthDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(5) = iIndex & ",ChildBirthDate"
				End Select
			End If
		Next

		sFields = "EmployeeID, ChildID, ChildName, ChildLastName, ChildLastName2, ChildBirthDate, ChildEndDate, LevelID, RegistrationDate, UserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "ChildName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nombre del hijo(a)."
		ElseIf InStr(1, sFields, "ChildLastName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el apellido paterno del hijo(a)."
		ElseIf InStr(1, sFields, "ChildBirthDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de nacimiento de los hijos."
		ElseIf InStr(1, sFields, "LevelID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nivel escolar de la beca."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "ChildBirthDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "ChildName"
								aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ChildLastName"
								aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ChildLastName2"
								aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "LevelID"
								aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case Else
						End Select
					Next
					lErrorNumber = SaveEmployeeChildren(aEmployeeComponent, "ChildrenSchoolarships", sErrorDescription)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadChildrenSchoolarshipsFile = lErrorNumber
	Err.Clear
End Function

Function UploadConceptsValuesFile(oADODBConnection, sFileName, bFull, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         ConceptValues table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadConceptsValuesFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iTotalRecord
	Dim sCompanyShortName
	Dim sConceptShortName
	Dim sLevelShortName
	Dim sPositionShortName
	Dim sPositionTypeShortName
	Dim sGroupGradeLevelShortName
	Dim sEmployeeTypeShortName
	Dim sStatusName
	Dim sJobStatusShortName
	Dim sJourneyShortName
	Dim sServiceShortName
	Dim sAntiquityName
	Dim sAntiquity2Name
	Dim sAntiquity3Name
	Dim sAntiquity4Name
	Dim sGenderName
	Dim sSchoolarshipName
	Dim sConceptCurrencyID
	Dim sAppliesToID
	Dim asAppliesToID
	Dim sConceptMinQttyID
	Dim sConceptMaxQttyID
	Dim lStartDateForValueConcept
	Dim lEndDateForValueConcept
	iTotalRecord = 0

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "CompanyID"
						aiFieldsOrder(iIndex) = iIndex & ",CompanyID"
					Case "ConceptID"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptID"
					Case "PositionID"
						aiFieldsOrder(iIndex) = iIndex & ",PositionID"
					Case "ConceptAmount"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptAmount"
					Case "OcurredStartDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "OcurredStartDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "OcurredStartDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "OcurredEndDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "OcurredEndDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "OcurredEndDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EconomicZoneID"
						aiFieldsOrder(iIndex) = iIndex & ",EconomicZoneID"
					Case "GroupGradeLevelID"
						aiFieldsOrder(iIndex) = iIndex & ",GroupGradeLevelID"
					Case "ClassificationID"
						aiFieldsOrder(iIndex) = iIndex & ",ClassificationID"
					Case "IntegrationID"
						aiFieldsOrder(iIndex) = iIndex & ",IntegrationID"
					Case "LevelID"
						aiFieldsOrder(iIndex) = iIndex & ",LevelID"
					Case "PositionShortNames"
						aiFieldsOrder(iIndex) = iIndex & ",PositionShortNames"
					Case "PositionTypeID"
						aiFieldsOrder(iIndex) = iIndex & ",PositionTypeID"
					Case "WorkingHours"
						aiFieldsOrder(iIndex) = iIndex & ",WorkingHours"
					Case "EmployeeTypeID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeTypeID"
					Case "EmployeeStatusID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeStatusID"
					Case "JobStatusID"
						aiFieldsOrder(iIndex) = iIndex & ",JobStatusID"
					Case "JobStatusID"
						aiFieldsOrder(iIndex) = iIndex & ",JobStatusID"
					Case "JourneyID"
						aiFieldsOrder(iIndex) = iIndex & ",JourneyID"
					Case "AdditionalShift"
						aiFieldsOrder(iIndex) = iIndex & ",AdditionalShift"
					Case "ServiceID"
						aiFieldsOrder(iIndex) = iIndex & ",ServiceID"
					Case "ForRisk"
						aiFieldsOrder(iIndex) = iIndex & ",ForRisk"
					Case "HasChildren"
						aiFieldsOrder(iIndex) = iIndex & ",HasChildren"
					Case "HasSyndicate"
						aiFieldsOrder(iIndex) = iIndex & ",HasSyndicate"
					Case "AntiquityID"
						aiFieldsOrder(iIndex) = iIndex & ",AntiquityID"
					Case "Antiquity2ID"
						aiFieldsOrder(iIndex) = iIndex & ",Antiquity2ID"
					Case "Antiquity3ID"
						aiFieldsOrder(iIndex) = iIndex & ",Antiquity3ID"
					Case "Antiquity4ID"
						aiFieldsOrder(iIndex) = iIndex & ",Antiquity4ID"
					Case "GenderID"
						aiFieldsOrder(iIndex) = iIndex & ",GenderID"
					Case "SchoolarshipID"
						aiFieldsOrder(iIndex) = iIndex & ",SchoolarshipID"
					Case "ConceptCurrencyID"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptCurrencyID"
					Case "AppliesToID"
						aiFieldsOrder(iIndex) = iIndex & ",AppliesToID"
					Case "ConceptMin"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptMin"
					Case "ConceptMinQttyID"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptMinQttyID"
					Case "ConceptMax"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptMax"
					Case "ConceptMaxQttyID"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptMaxQttyID"
				End Select
			End If
		Next
		If Not bFull Then
			If CInt(oRequest("EmployeeTypeID").Item) = -1 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene el tipo de tabulador."
			Else
				aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CInt(oRequest("EmployeeTypeID").Item)
				Select Case aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)
					Case 0
						sFields = ",CompanyID;,ConceptID;,PositionShortNames;,ConceptAmount;,StartDate;,PositionTypeID;,LevelID;,WorkingHours;,EconomicZoneID;"
					Case 1
						sFields = ",CompanyID;,ConceptID;,PositionShortNames;,ConceptAmount;,StartDate;,GroupGradeLevelID;,ClassificationID;,IntegrationID;"
					Case 2
						sFields = ",CompanyID;,ConceptID;,PositionShortNames;,ConceptAmount;,StartDate;,PositionTypeID;,LevelID;,EconomicZoneID;"
					Case 3
						sFields = ",CompanyID;,ConceptID;,PositionShortNames;,ConceptAmount;,StartDate;,LevelID;"
					Case 4, 5, 6
						sFields = ",CompanyID;,ConceptID;,PositionShortNames;,ConceptAmount;,StartDate;,LevelID;,EconomicZoneID;"
				End Select
			End If
		Else
			sFields = ",EmployeeTypeID;,CompanyID;,ConceptID;,PositionTypeID;,PositionShortNames;,LevelID;,GroupGradeLevelID;,EmployeeStatusID;,JobStatusID;,ClassificationID;,IntegrationID;,JourneyID;,WorkingHours;,AdditionalShift;,EconomicZoneID;,ServiceID;,AntiquityID;,Antiquity2ID;,Antiquity3ID;,Antiquity4ID;,ForRisk;,GenderID;,HasChildren;,SchoolarshipID;,HasSyndicate;,ConceptAmount;,StartDate;,ConceptCurrencyID;,AppliesToID;,ConceptMin;,ConceptMinQttyID;,ConceptMax;,ConceptMaxQttyID;"
		End If
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, ("," & aiFieldsOrder(iIndex)(1) & ";"), "")
		Next
		If (lErrorNumber <> -1) And (InStr(1, sFields, "CompanyID") > 0) Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la compañía."
		ElseIf InStr(1, sFields, "ConceptID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la clave del concepto de pago."
		ElseIf InStr(1, sFields, "PositionShortNames") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el puesto."
		ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el monto a pagar."
		ElseIf InStr(1, sFields, "StartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de inicio de la vigencia."
		ElseIf InStr(1, sFields, "PositionTypeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el tipo de puesto."
		ElseIf InStr(1, sFields, "LevelID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nivel del puesto."
		ElseIf InStr(1, sFields, "WorkingHours") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las horas laboradas(Jornada)."
		ElseIf InStr(1, sFields, "EconomicZoneID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la zona económica."
		ElseIf InStr(1, sFields, "GroupGradeLevelID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Grupo grado nivel del puesto."
		ElseIf InStr(1, sFields, "ClassificationID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la clasificación del puesto."
		ElseIf InStr(1, sFields, "IntegrationID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la integración del puesto."
		ElseIf InStr(1, sFields, "EmployeeTypeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Tipo de tabulador al que aplicara el registro."
		ElseIf InStr(1, sFields, "EmployeeStatusID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el estatus del empleado al que aplicara el registro."
		ElseIf InStr(1, sFields, "JobStatusID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el estatus de la plaza al que aplicara el registro."
		ElseIf InStr(1, sFields, "JourneyID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Jornada para el tabulador."
		ElseIf InStr(1, sFields, "AdditionalShift") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no indica si el tabulador aplica para Turno opcional (0 - NO; 1 - SI)."
		ElseIf InStr(1, sFields, "ServiceID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el servicio para el tabulador."
		ElseIf InStr(1, sFields, "ForRisk") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no indica si el tabulador aplica para Riesgos profesionales (0 - NO; 1 - SI)."
		ElseIf InStr(1, sFields, "HasChildren") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no indica si el tabulador aplica para empleados con hijos (0 - NO; 1 - SI)."
		ElseIf InStr(1, sFields, "HasSyndicate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no indica si el tabulador aplica para empleados sindicalizados (0 - NO; 1 - SI)."
		ElseIf InStr(1, sFields, "AntiquityID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Antigüedad en el ISSSTE para el tabulador."
		ElseIf InStr(1, sFields, "Antiquity2ID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Antigüedad consecutiva para el tabulador."
		ElseIf InStr(1, sFields, "Antiquity3ID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Antigüedad en el ISSSTE con plaza de base para el tabulador."
		ElseIf InStr(1, sFields, "Antiquity4ID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Antigüedad federal para el tabulador."
		ElseIf InStr(1, sFields, "GenderID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el género del empleado para el tabulador."
		ElseIf InStr(1, sFields, "SchoolarshipID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la escolaridad de los hijos del empleado para el tabulador."
		ElseIf InStr(1, sFields, "ConceptCurrencyID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la unidad del monto del para el tabulador."
		ElseIf InStr(1, sFields, "AppliesToID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene los conceptos sobre los que aplica (separados por comas) el tabulador."
		ElseIf InStr(1, sFields, "ConceptMin") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el monto mínimo a pagar."
		ElseIf InStr(1, sFields, "ConceptMinQttyID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la unidad del monto del para el tabulador."
		ElseIf InStr(1, sFields, "ConceptMax") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el monto máximo pagar."
		ElseIf InStr(1, sFields, "ConceptMaxQttyID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la unidad del monto del para el tabulador."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					Err.Clear
					lErrorNumber = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					sCompanyShortName = " Ninguna"
					sLevelShortName = " N"
					sGroupGradeLevelShortName = "N"
					sPositionTypeShortName = " N"
					sEmployeeTypeShortName = " N"
					sStatusName = "-1"
					sJobStatusShortName = " N"
					sJourneyShortName = " N"
					sServiceShortName = "00000"
					sAntiquityName = " Ninguna"
					sAntiquity2Name = " Ninguna"
					sAntiquity3Name = " Ninguna"
					sAntiquity4Name = " Ninguna"
					sGenderName = "Ninguno"
					sSchoolarshipName = " Ninguna"
					sConceptCurrencyID = "$"
					sAppliesToID = "-1"
					sConceptMinQttyID = "$"
					sConceptMaxQttyID = "$"
					aConceptComponent(N_COMPANY_ID_CONCEPT) = -1
					aConceptComponent(N_RECORD_ID_CONCEPT) = -1
					aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = -1
					aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = -1
					aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = -1
					aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
					aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
					aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
					aConceptComponent(N_JOURNEY_ID_CONCEPT) = -1
					aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1
					aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
					aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
					aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0
					aConceptComponent(N_SERVICE_ID_CONCEPT) = -1
					aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1
					aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = -1
					aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = -1
					aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = -1
					aConceptComponent(N_FOR_RISK_CONCEPT) = -1
					aConceptComponent(N_GENDER_ID_CONCEPT) = -1
					aConceptComponent(N_HAS_CHILDREN_CONCEPT) = -1
					aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = -1
					aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = -1
					aConceptComponent(N_CURRENCY_ID_CONCEPT) = 0
					aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = 1
					aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = 1
					aConceptComponent(S_APPLIES_ID_CONCEPT) = -1
					aConceptComponent(D_CONCEPT_MIN_CONCEPT) = 0
					aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) = 1
					aConceptComponent(D_CONCEPT_MAX_CONCEPT) = 0
					aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) = 1
					aConceptComponent(N_START_USER_ID_CONCEPT) = aLoginComponent(N_USER_ID_LOGIN)
					aConceptComponent(N_STATUS_ID_CONCEPT) = 0
					aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
					For jIndex = 0 To UBound(aiFieldsOrder)
						Err.Clear
						Select Case aiFieldsOrder(jIndex)(1)
							Case "CompanyID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sCompanyShortName = " Ninguna"
								Else
									sCompanyShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									If sCompanyShortName = "ISSSTE" Then sCompanyShortName = "ISSSTE-ASEGURADOR"
								End If
							Case "ClassificationID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
								Else
									aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								End If
							Case "ConceptAmount"
								aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								If CInt(Request.Cookies("SIAP_SectionID")) <> 4 Then
									aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT)/2)
								End If
							Case "ConceptID"
								sConceptShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EconomicZoneID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0
								Else
									aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								End If
							Case "GroupGradeLevelID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sGroupGradeLevelShortName = "N"
								Else
									sGroupGradeLevelShortName = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								End If
							Case "IntegrationID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
								Else
									aConceptComponent(N_INTEGRATION_ID_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								End If
							Case "LevelID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sLevelShortName = " N"
								Else
									sLevelShortName = Right("000" & Trim(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), Len("000"))
								End If
							Case "PositionShortNames"
								sPositionShortName = CStr(Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´"))
								aConceptComponent(S_POSITION_SHORT_NAME_CONCEPT) = sPositionShortName
							Case "PositionTypeID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sPositionTypeShortName = " N"
								Else
									sPositionTypeShortName = CStr(Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´"))
								End If
							Case "WorkingHours"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1
								Else
									aConceptComponent(D_WORKING_HOURS_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								End If
							Case "StartDate"
								If lErrorNumber = 0 Then
									Select Case sDateFormat
										Case "YYYYMMDD"
											lStartDateForValueConcept = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											lStartDateForValueConcept = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											lStartDateForValueConcept = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(lStartDateForValueConcept) Then
										aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(lStartDateForValueConcept)
									Else
										Err.Clear
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "EndDate"
								If lErrorNumber = 0 Then
									Select Case sDateFormat
										Case "YYYYMMDD"
											lEndDateForValueConcept = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											lEndDateForValueConcept = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											lEndDateForValueConcept = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(lEndDateForValueConcept) Then
										aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = CLng(lEndDateForValueConcept)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "EmployeeTypeID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sEmployeeTypeShortName = " N"
								Else
									sEmployeeTypeShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "JourneyID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sJourneyShortName = " N"
								Else
									sJourneyShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "EmployeeStatusID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sStatusName = "-1"
								Else
									sStatusName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "JobStatusID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sJobStatusShortName = " N"
								Else
									sJobStatusShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "AdditionalShift"
								If UCase(CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))) = "SI" Then
									aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = 1
								Else
									aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = 0
								End If
							Case "ServiceID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sServiceShortName = "00000"
								Else
									sServiceShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "AntiquityID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sAntiquityName = " Ninguna"
								Else
									sAntiquityName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "Antiquity2ID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sAntiquity2Name = " Ninguna"
								Else
									sAntiquity2Name = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "Antiquity3ID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sAntiquity3Name = " Ninguna"
								Else
									sAntiquity3Name = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "Antiquity4ID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sAntiquity4Name = " Ninguna"
								Else
									sAntiquity4Name = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "ForRisk"
								If UCase(CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))) = "SI" Then
									aConceptComponent(N_FOR_RISK_CONCEPT) = 1
								Else
									aConceptComponent(N_FOR_RISK_CONCEPT) = 0
								End If
							Case "GenderID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sGenderName = "Ninguno"
								Else
									sGenderName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "HasChildren"
								If UCase(CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))) = "SI" Then
									aConceptComponent(N_HAS_CHILDREN_CONCEPT) = 1
								Else
									aConceptComponent(N_HAS_CHILDREN_CONCEPT) = 0
								End If
							Case "SchoolarshipID"
								If (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todos") > 0) Or (InStr(1, CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))), "Todas") > 0) Then
									sSchoolarshipName = " Ninguna"
								Else
									sSchoolarshipName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								End If
							Case "HasSyndicate"
								If UCase(CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))) = "SI" Then
									aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = 1
								Else
									aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = 0
								End If
							Case "ConceptCurrencyID"
								sConceptCurrencyID = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "AppliesToID"
								sAppliesToID = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								asAppliesToID = Split(sAppliesToID, ",")
								sAppliesToID = ""
								For kIndex = 0 To UBound(asAppliesToID)
									If VerifyExistenceOfCatalogInDate(oADODBConnection, "Concepts", "ConceptShortName,StartDate,EndDate", CStr(asAppliesToID(kIndex) & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
										sAppliesToID = sAppliesToID & CStr(oRecordset.Fields("ConceptID").Value) & ","
									Else
										lErrorNumber = -1
										sErrorDescription = "El concepto " & asAppliesToID(kIndex) & " sobre el que se aplica no existe en el catálogo."
									End If
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
								Next
								If lErrorNumber = 0 Then
									If (InStr(Right(sAppliesToID, 1), ",") > 0) Then
										sAppliesToID = Left(sAppliesToID, Len(sAppliesToID) -1)
									End If
								End If
							Case "ConceptMin"
								aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "ConceptMinQttyID"
								sConceptMinQttyID = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "ConceptMax"
								aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "ConceptMaxQttyID"
								sConceptMaxQttyID = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case Else
						End Select
					Next
					If lErrorNumber = 0 Then
						aConceptComponent(N_START_DATE_CONCEPT) = aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT)
						aConceptComponent(N_END_DATE_CONCEPT) = aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT)
						If VerifyExistenceOfCatalogInDate(oADODBConnection, "Companies", "CompanyName,StartDate,EndDate", CStr(sCompanyShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
							aConceptComponent(N_COMPANY_ID_CONCEPT) = CLng(oRecordset.Fields("CompanyID").Value)
							If VerifyExistenceOfCatalogInDate(oADODBConnection, "Concepts", "ConceptShortName,StartDate,EndDate", CStr(sConceptShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
								aConceptComponent(N_ID_CONCEPT) = CLng(oRecordset.Fields("ConceptID").Value)
								If VerifyExistenceOfCatalogInDate(oADODBConnection, "Levels", "LevelShortName,StartDate,EndDate", CStr(sLevelShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
									aConceptComponent(N_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("LevelID").Value)
									If VerifyExistenceOfCatalogInDate(oADODBConnection, "PositionTypes", "PositionTypeShortName,StartDate,EndDate", CStr(sPositionTypeShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
										If VerifyExistenceOfCatalogInDate(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelShortName,StartDate,EndDate", CStr(sGroupGradeLevelShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
											aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
											If Not bFull Then
												If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 1 Then aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
												lErrorNumber = CheckExistencyOfPosition(aConceptComponent, sPositionShortName, sErrorDescription)
												If lErrorNumber = 0 Then
													lErrorNumber = AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
												Else
													lErrorNumber = L_ERR_NO_RECORDS
													sErrorDescription = "El puesto " & sPositionShortName & " para la compañía " & sCompanyShortName  & " no existe en la base de datos."
												End If
											Else
												If VerifyExistenceOfCatalogInDate(oADODBConnection, "EmployeeTypes", "EmployeeTypeShortName,StartDate,EndDate", CStr(sEmployeeTypeShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
													aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
													If VerifyExistenceOfCatalogInDate(oADODBConnection, "StatusEmployees", "StatusID", CStr(sStatusName), "", oRecordset, sErrorDescription) Then
														aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = CLng(oRecordset.Fields("StatusID").Value)
														If VerifyExistenceOfCatalogInDate(oADODBConnection, "StatusJobs", "StatusShortName", CStr(sJobStatusShortName), "", oRecordset, sErrorDescription) Then
															aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = CLng(oRecordset.Fields("StatusID").Value)
															If VerifyExistenceOfCatalogInDate(oADODBConnection, "Journeys", "JourneyShortName,StartDate,EndDate", CStr(sJourneyShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
																aConceptComponent(N_JOURNEY_ID_CONCEPT) = CLng(oRecordset.Fields("JourneyID").Value)
																If VerifyExistenceOfCatalogInDate(oADODBConnection, "Services", "ServiceShortName,StartDate,EndDate", CStr(sServiceShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
																	aConceptComponent(N_SERVICE_ID_CONCEPT) = CLng(oRecordset.Fields("ServiceID").Value)
																	If VerifyExistenceOfCatalogInDate(oADODBConnection, "Antiquities", "AntiquityName", CStr(sAntiquityName), "", oRecordset, sErrorDescription) Then
																		aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = CLng(oRecordset.Fields("AntiquityID").Value)
																		If VerifyExistenceOfCatalogInDate(oADODBConnection, "Antiquities", "AntiquityName", CStr(sAntiquity2Name), "", oRecordset, sErrorDescription) Then
																			aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = CLng(oRecordset.Fields("AntiquityID").Value)
																			If VerifyExistenceOfCatalogInDate(oADODBConnection, "Antiquities", "AntiquityName", CStr(sAntiquity3Name), "", oRecordset, sErrorDescription) Then
																				aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = CLng(oRecordset.Fields("AntiquityID").Value)
																				If VerifyExistenceOfCatalogInDate(oADODBConnection, "Antiquities", "AntiquityName", CStr(sAntiquity4Name), "", oRecordset, sErrorDescription) Then
																					aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = CLng(oRecordset.Fields("AntiquityID").Value)
																					If VerifyExistenceOfCatalogInDate(oADODBConnection, "Genders", "GenderName", CStr(sGenderName), "", oRecordset, sErrorDescription) Then
																						aConceptComponent(N_GENDER_ID_CONCEPT) = CLng(oRecordset.Fields("GenderID").Value)
																						If VerifyExistenceOfCatalogInDate(oADODBConnection, "Schoolarships", "SchoolarshipName", CStr(sSchoolarshipName), "", oRecordset, sErrorDescription) Then
																							aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = CLng(oRecordset.Fields("SchoolarshipID").Value)
																							lErrorNumber = CheckExistencyOfPosition(aConceptComponent, sPositionShortName, sErrorDescription)
																							If lErrorNumber = 0 Then
																								lErrorNumber = AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
																							Else
																								lErrorNumber = L_ERR_NO_RECORDS
																								sErrorDescription = "El puesto " & sPositionShortName & " para la compañía " & sCompanyShortName  & " no existe en la base de datos."
																							End If
																						Else
																							lErrorNumber = -1
																							sErrorDescription = "No esta registrado alguna escolaridad con la clave indicada."
																						End If
																					Else
																						lErrorNumber = -1
																						sErrorDescription = "No esta registrado algún genero con la clave indicada."
																					End If
																				Else
																					lErrorNumber = -1
																					sErrorDescription = "No esta registrada alguna Antigüedad federal con la clave indicada: " & sAntiquity4Name
																				End If
																			Else
																				lErrorNumber = -1
																				sErrorDescription = "No esta registrada alguna Antigüedad en el ISSSTE con plaza de base con la clave indicada: " & sAntiquity3Name
																			End If
																		Else
																			lErrorNumber = -1
																			sErrorDescription = "No esta registrada alguna Antigüedad consecutiva con la clave indicada: " & sAntiquity2Name
																		End If
																	Else
																		lErrorNumber = -1
																		sErrorDescription = "No esta registrado algún Antigüedad en el ISSSTE con la clave indicada: " & sAntiquityName
																	End If
																Else
																	lErrorNumber = -1
																	sErrorDescription = "No esta registrado algún servicio con la clave indicada."
																End If
															Else
																lErrorNumber = -1
																sErrorDescription = "No esta registrado alguna Jornada con la clave indicada."
															End If
														Else
															lErrorNumber = -1
															sErrorDescription = "No esta registrado algún estatus de la plaza con la clave indicada."
														End If
													Else
														lErrorNumber = -1
														sErrorDescription = "No esta registrado algún estatus del empleado con la clave indicada."
													End If
												Else
													lErrorNumber = -1
													sErrorDescription = "No esta registrado algún tipo de tabulador con la clave indicada."
												End If
											End If
										Else
											lErrorNumber = -1
											sErrorDescription = "No esta registrado algún grupo-grado-nivel con la clave indicada."
										End If
									Else
										lErrorNumber = -1
										sErrorDescription = "No esta registrado el tipo de puesto indicado."
									End If
								Else
									lErrorNumber = -1
									sErrorDescription = "No esta registrado el nivel indicado."
								End If
							Else
								lErrorNumber = -1
								sErrorDescription = "No esta registrado algún concepto con la clave indicada."
							End If
						Else
							lErrorNumber = -1
							sErrorDescription = "No esta registrada alguna compañía con la clave indicada."
						End If
						If lErrorNumber <> 0 Then
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						Else
							iTotalRecord = iTotalRecord + 1
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If iTotalRecord > 0 Then
			Call DisplayErrorMessage("Confirmación", "Han sido registrados " & iTotalRecord & " renglones del tabulador cargado.")
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadConceptsValuesFile = lErrorNumber
	Err.Clear
End Function

Function UploadConsarFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         DM_APORT_SAR table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadConsarFile"
	Dim sFileContents
	Dim asFileContents
	Dim iIndex
	Dim jIndex
	Dim sErrorQueries
	Dim lErrorNumber
	Dim asFileRow
	Dim sQuery
	Dim sDate
	Dim lCurrentDate
	Dim iEmployeeID
	Dim iCT
	Dim iCompanyID
	Dim iZoneCode
	Dim sRfc
	Dim sCurp
	Dim sEmployeeData
	Dim oRecordset

    Dim sDuplicatedRfc
    Dim sDuplicatedCurp
    Dim sEndPayroll
    
	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	sQuery = "Truncate Table DM_APORT_SAR"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	sQuery = "Insert Into DM_APORT_SAR ("
	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		sErrorQueries = ""
		For iIndex = 0 To UBound(asFileContents)
			sQuery = "Insert Into DM_APORT_SAR ("
			asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
			For jIndex = 0 To UBound(asFileRow)
				If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
					If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
						If jIndex < UBound(asFileRow) Then
							sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ","
						Else
							sQuery = sQuery & sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ") Values ("'",EmployeeID,CT,ZoneCode,CompanyID) Values ("
						End If
					Else
						If jIndex < UBound(asFileRow) Then
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ","
						Else
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ") Values ("'",EmployeeID,CT,ZoneCode,CompanyID) Values ("
						End If
					End If
				End If
			Next

			For jIndex = 0 To UBound(asFileRow)
				If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
					If (strComp(oRequest("Column" & (jIndex + 1)),"rfc",vbBinaryCompare) = 0) Then sRFC = asFileRow(jIndex)
					If (strComp(oRequest("Column" & (jIndex + 1)),"curp",vbBinaryCompare) = 0) Then sCURP = asFileRow(jIndex)
					If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Then sDate = asFileRow(jIndex)
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),1,2) & Mid(asFileRow(jIndex),4,2)
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),4,2) & Mid(asFileRow(jIndex),1,2)
							If Len(sDate) = 0 Then sDate = "0"
							sQuery = sQuery & sDate & ","
					Else
						If isNumeric(asFileRow(jIndex)) Then
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & asFileRow(jIndex) & ","
							Else
								sQuery = sQuery & asFileRow(jIndex) & ")"
							End If
						Else
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "',"
							Else
								sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "')"
							End If
						End If
					End If
				End If
			Next
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		Next

        'Buscamos en la tabla dm_padron_banamex los rfc duplicados
        sDuplicatedRfc=""
        sQuery="select rfc, count(*) from dm_padron_banamex group by rfc having count(*)>1"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
        Do While Not oRecordset.EOF
           sDuplicatedRfc= sDuplicatedRfc & "'" & oRecordset.Fields("rfc").Value & "',"
           oRecordset.MoveNext
        Loop
        If sDuplicatedRfc <> "" Then 
           sDuplicatedRfc= left(sDuplicatedRfc, len(sDuplicatedRfc)-1)
           sDuplicatedRfc=" AND rfc not in (" & sDuplicatedRfc & ")"
        End If
		oRecordset.Close

        'Buscamos en la tabla dm_padron_banamex los curp duplicados
        sDuplicatedCurp=""
        sQuery="select curp, count(*) from dm_padron_banamex group by curp having count(*)>1"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
        Do While Not oRecordset.EOF
           sDuplicatedCurp= sDuplicatedCurp & "'" & oRecordset.Fields("curp").Value & "',"
           oRecordset.MoveNext
        Loop
        If sDuplicatedCurp <> "" Then 
           sDuplicatedCurp= left(sDuplicatedCurp, len(sDuplicatedCurp)-1)
           sDuplicatedCurp=" AND curp not in (" & sDuplicatedCurp & ")"
        End If
		oRecordset.Close

        sErrorDescription = sErrorQueries
		sQuery = "Update DM_APORT_SAR Set Salary = Salary / 100, SalaryV = SalaryV / 100, sar = sar / 100, EntityCV = EntityCV / 100, EmployeeCV = EmployeeCV / 100, foviAmount = foviAmount / 100, EmployeeSaving = EmployeeSaving / 100, entitySaving = entitySaving / 100"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If iConnectionType = ORACLE Then
			'sQuery = "Update siap.dm_aport_sar sar Set (employeeid, companyid, ct) = (Select employeeid, u_version, ct From siap.dm_padron_banamex bx Where (sar.rfc = bx.rfc))"
            sQuery = "Update siap.dm_aport_sar sar Set (employeeid, companyid, ct) = (Select employeeid, u_version, ct From siap.dm_padron_banamex bx Where (sar.rfc = bx.rfc)" & sDuplicateRfc & "  )" 
		Else
			'sQuery = "Update dm_aport_sar Set employeeid = dm_padron_banamex.employeeid,companyid = dm_padron_banamex.u_version,ct = dm_padron_banamex.ct From dm_padron_banamex Where (dm_aport_sar.rfc = dm_padron_banamex.rfc)"
            sQuery = "Update dm_aport_sar Set employeeid = dm_padron_banamex.employeeid,companyid = dm_padron_banamex.u_version,ct = dm_padron_banamex.ct From dm_padron_banamex Where (dm_aport_sar.rfc = dm_padron_banamex.rfc)" & sDuplicatedRfc
		End If
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If iConnectionType = ORACLE Then
			'sQuery = "Update siap.dm_aport_sar sar Set (employeeid, companyid, ct) = (Select employeeid, u_version, ct From siap.dm_padron_banamex bx Where (sar.curp = bx.curp))"
            sQuery = "Update siap.dm_aport_sar sar Set (employeeid, companyid, ct) = (Select employeeid, u_version, ct From siap.dm_padron_banamex bx Where (sar.curp = bx.curp)" & sDuplicatedCurp & " )"
		Else
			'sQuery = "Update dm_aport_sar Set employeeid = dm_padron_banamex.employeeid,companyid = dm_padron_banamex.u_version,ct = dm_padron_banamex.ct From dm_padron_banamex Where (dm_aport_sar.curp = dm_padron_banamex.curp)"
            sQuery = "Update dm_aport_sar Set employeeid = dm_padron_banamex.employeeid,companyid = dm_padron_banamex.u_version,ct = dm_padron_banamex.ct From dm_padron_banamex Where (dm_aport_sar.curp = dm_padron_banamex.curp) " & sDuplicatedCurp
		End If
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

        sEndPayroll=0
        sQuery="select endpayroll from dm_sar_periods where isopen=1"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
       If Not oRecordset.EOF Then sEndPayroll= Cint(oRecordset.Fields("endpayroll").Value)

        sQuery = "Update dm_aport_sar Set zonecode =(select zoneid from employeeshistorylistforpayroll where payrollid=" & sEndPayroll & " and dm_aport_sar.employeeid=employeeshistorylistforpayroll.employeeid)"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	End If
End Function

Function UploadDocumentsForLicensesFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         DocumentsForLicenses table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadDocumentsForLicensesFile"
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormatDocument
	Dim sDateFormatStart
	Dim sDateFormatEnd
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iDocumentDateYear
	Dim iDocumentDateMonth

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "DocumentForLicenseNumber"
						aiFieldsOrder(0) = iIndex & ",DocumentForLicenseNumber"
					Case "DocumentTemplate"
						aiFieldsOrder(1) = iIndex & ",DocumentTemplate"
					Case "RequestNumber"
						aiFieldsOrder(2) = iIndex & ",RequestNumber"
					Case "EmployeeID"
						aiFieldsOrder(3) = iIndex & ",EmployeeID"
					Case "LicenseSyndicateTypeID"
						aiFieldsOrder(4) = iIndex & ",LicenseSyndicateTypeID"
					Case "OcurredDocumentDateYYYYMMDD"
						sDateFormatDocument = "YYYYMMDD"
						aiFieldsOrder(5) = iIndex & ",DocumentLicenseDate"
					Case "OcurredDocumentDateDDMMYYYY"
						sDateFormatDocument = "DDMMYYYY"
						aiFieldsOrder(5) = iIndex & ",DocumentLicenseDate"
					Case "OcurredDocumentDateMMDDYYYY"
						sDateFormatDocument = "MMDDYYYY"
						aiFieldsOrder(5) = iIndex & ",DocumentLicenseDate"
					Case "OcurredStartDateYYYYMMDD"
						sDateFormatStart = "YYYYMMDD"
						aiFieldsOrder(6) = iIndex & ",LicenseStartDate"
					Case "OcurredStartDateDDMMYYYY"
						sDateFormatStart = "DDMMYYYY"
						aiFieldsOrder(6) = iIndex & ",LicenseStartDate"
					Case "OcurredStartDateMMDDYYYY"
						sDateFormatStart = "MMDDYYYY"
						aiFieldsOrder(6) = iIndex & ",LicenseStartDate"
					Case "OcurredEndDateYYYYMMDD"
						sDateFormatEnd = "YYYYMMDD"
						aiFieldsOrder(7) = iIndex & ",LicenseEndDate"
					Case "OcurredEndDateDDMMYYYY"
						sDateFormatEnd = "DDMMYYYY"
						aiFieldsOrder(7) = iIndex & ",LicenseEndDate"
					Case "OcurredEndDateMMDDYYYY"
						sDateFormatEnd = "MMDDYYYY"
						aiFieldsOrder(7) = iIndex & ",LicenseEndDate"
				End Select
			End If
		Next

		sFields = "DocumentForLicenseID, DocumentForLicenseNumber, DocumentTemplate, RequestNumber, EmployeeID, LicenseSyndicateTypeID, DocumentLicenseDate, LicenseStartDate, LicenseEndDate, DocumentForCancelLicenseNumber, LicenseCancelDate, UserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "DocumentForLicenseNumber") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número del oficio."
		ElseIf InStr(1, sFields, "DocumentTemplate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nombre de la plantilla."
		ElseIf InStr(1, sFields, "RequestNumber") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de la solicitud."
		ElseIf InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "LicenseSyndicateTypeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el tipo de licencia sindical."
		ElseIf InStr(1, sFields, "DocumentLicenseDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha del oficio."
		ElseIf InStr(1, sFields, "LicenseStartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de inicio de la licencia."
		ElseIf InStr(1, sFields, "LicenseEndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de término de la licencia."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sValues = Replace(Replace(Replace(sFields, "UserID", aLoginComponent(N_USER_ID_LOGIN)),"DocumentForCancelLicenseNumber",0),"LicenseCancelDate",0)
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)

					sQuery = "Insert Into DocumentsForLicenses ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "DocumentLicenseDate"
								Select Case sDateFormatDocument
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
										iDocumentDateYear = CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0))) * .0001)
										iDocumentDateMonth = (asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) - (iDocumentDateYear * 10000)
										iDocumentDateMonth = CInt(iDocumentDateMonth / 100)
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
										iDocumentDateYear = asInputDate(2)
										iDocumentDateMonth = asInputDate(1)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
										iDocumentDateYear = asInputDate(2)
										iDocumentDateMonth = asInputDate(0)
								End Select
							Case "LicenseStartDate"
								Select Case sDateFormatStart
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case "LicenseEndDate"
								Select Case sDateFormatEnd
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case "DocumentForLicenseNumber", "DocumentTemplate", "RequestNumber"
								sQuery = sQuery & "'" & Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´") & "', "
							Case ""
							Case Else
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
						End Select
					Next
					If InStr(1, sValues, "DocumentDateYear") > 0 Then
						sValues = Replace(sValues, "DocumentDateYear", iDocumentDateYear)
					End If
					If InStr(1, sValues, "DocumentDateMonth") > 0 Then
						sValues = Replace(sValues, "DocumentDateMonth", iDocumentDateMonth)
					End If
					sErrorDescription = "No se pudo obtener un identificador para el nuevo documento."
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "DocumentsForLicenses", "DocumentForLicenseID", "", 1, aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE), sErrorDescription)
					If InStr(1, sValues, "DocumentForLicenseID") > 0 Then
						sExecuteQuery = Replace(sValues, "DocumentForLicenseID", aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE))
					End If
					sQuery = sQuery & sExecuteQuery & ")"
					sErrorDescription = "No se pudo guardar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadDocumentsForLicensesFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesAbsencesFile(oADODBConnection, lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesAbsencesLKP table.
'Inputs:  oADODBConnection, lReasonID, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesAbsencesFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim sEndDateFormat
	Dim sPayrollDateFormat
	Dim sStartDate
	Dim sEndDate
	Dim sPayrollDate
	Dim asInputDate
	Dim sFields
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iEmployeeTypeID
	Dim iPositionTypeID
	Dim sPeriod
	Dim sPeriodYear
	Dim sAbsenceIDs
	Dim sAbsenceType
	Dim iAbsenceID
	Dim iActiveOriginal
	Dim sAbsenceShortName
	Dim sJustificationShortName

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeID"
					Case "AbsenceID"
						aiFieldsOrder(iIndex) = iIndex & ",AbsenceID"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "EndDateYYYYMMDD"
						sEndDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateDDMMYYYY"
						sEndDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateMMDDYYYY"
						sEndDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "VacationPeriod"
						aiFieldsOrder(iIndex) = iIndex & ",VacationPeriod"
					Case "PeriodYear"
						aiFieldsOrder(iIndex) = iIndex & ",PeriodYear"
					Case "Reasons"
						aiFieldsOrder(iIndex) = iIndex & ",Reasons"
					Case "DocumentNumber"
						aiFieldsOrder(iIndex) = iIndex & ",DocumentNumber"
					Case "AbsenceHours"
						aiFieldsOrder(iIndex) = iIndex & ",AbsenceHours"
					Case "JustificationID"
						aiFieldsOrder(iIndex) = iIndex & ",JustificationID"
					Case "AppliesForPunctuality"
						aiFieldsOrder(iIndex) = iIndex & ",AppliesForPunctuality"
					Case "ForJustificationID"
						aiFieldsOrder(iIndex) = iIndex & ",ForJustificationID"
				End Select
			End If
		Next
		Select Case lReasonID
			Case EMPLOYEES_EXTRAHOURS
				sFields = "EmployeeID, OcurredDate, PayrollDate, AbsenceHours, Reasons"
			Case EMPLOYEES_SUNDAYS
				sFields = "EmployeeID, OcurredDate, PayrollDate, Reasons"
			Case 0, 1
				sFields = "EmployeeID, AbsenceID, OcurredDate, PayrollDate"
			Case Else
				sFields = "EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate"
		End Select
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "AbsenceID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el tipo de incidencias."
		ElseIf InStr(1, sFields, "OcurredDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de las incidencias."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de termino de las incidencias."
		ElseIf (InStr(1, sFields, "AbsenceHours") > 0) And (lReasonID = EMPLOYEES_EXTRAHOURS) Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las horas extras."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					sEndDate = Empty
					sPayrollDate = Empty
					lErrorNumber = 0
					aAbsenceComponent(N_END_DATE_ABSENCE) = 0
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								sErrorDescription = "No existe el empleado indicado"
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									sErrorDescription = "Error al verificar la existencia del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID = '" & aEmployeeComponent(N_ID_EMPLOYEE) & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									iEmployeeTypeID = oRecordset.Fields("EmployeeTypeID").Value
									iPositionTypeID = oRecordset.Fields("PositionTypeID").Value
									oRecordset.Close
								End If
							Case "AbsenceID"
								If lErrorNumber = 0 Then
									sAbsenceShortName = Right(("0000" & CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))), Len("0000"))
									sQuery = "Select * from Absences Where (AbsenceShortName = '" & sAbsenceShortName & "')"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If oRecordset.EOF Then
											lErrorNumber = -1
											lConceptError = lConceptError + 1
											sErrorDescription = "No existe identificador para la clave de la incidencia indicada: " & sAbsenceShortName
										Else
											aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("AbsenceID").Value)
											aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
										End If
										oRecordset.Close
									End If
								End If
							Case "OcurredDate"
								If lErrorNumber = 0 Then
									Select Case sDateFormat
										Case "YYYYMMDD"
											sStartDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(sStartDate)
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "EndDate"
								If lErrorNumber = 0 Then
									Select Case sEndDateFormat
										Case "YYYYMMDD"
											sEndDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sEndDate) Then
										aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(sEndDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_END_DATE_ABSENCE) = aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "PayrollDate"
								If lErrorNumber = 0 Then
									Select Case sPayrollDateFormat
										Case "YYYYMMDD"
											sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sPayrollDate) Then
										aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CLng(sPayrollDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la quincena de aplicación en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aAbsenceComponent(N_APPLIED_DATE_ABSENCE), sErrorDescription) Then
											lErrorNumber = -1
										Else
											If (lReasonID = 0) Or (lReasonID = 1) Then
												If Not VerifyPayrollIsActive(oADODBConnection, aAbsenceComponent(N_APPLIED_DATE_ABSENCE), N_PAYROLL_FOR_ABSENCES, sErrorDescription) Then
													lErrorNumber = -1
												End If
											Else
												If Not VerifyPayrollIsActive(oADODBConnection, aAbsenceComponent(N_APPLIED_DATE_ABSENCE), N_PAYROLL_FOR_FEATURES, sErrorDescription) Then
													lErrorNumber = -1
												End If
											End If
										End If
									End If
								End If
							Case "DocumentNumber"
								aEmployeeComponent(xxxx) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "AbsenceHours"
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								Select Case lReasonID
									Case EMPLOYEES_EXTRAHOURS
										If lErrorNumber = 0 Then
											If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 1) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 3) Then
												sErrorDescription = "Solamente se pueden registrar 1, 2 o 3 horas extras en un día"
												lErrorNumber = -1
											End If
										End If
									Case EMPLOYEES_SUNDAYS
								End Select
							Case "JustificationID"
								aEmployeeComponent(N_CONCEPT_JUSTIFICATION_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "AppliesForPunctuality"
								aEmployeeComponent(N_CONCEPT_FOR_PUNCTUALITY_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "VacationPeriod"
								sPeriod = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If (Not IsEmpty(sPeriod)) And (Len(sPeriod) > 0) Then
									Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
										Case 35
											If (CInt(sPeriod) <> 1) And (CInt(sPeriod) <> 2) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1 y 2"
												lErrorNumber = -1
											End If
										Case 37
											If (CInt(sPeriod) <> 1) And (CInt(sPeriod) <> 2) And (CInt(sPeriod) <> 3) And (CInt(sPeriod) <> 4) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1, 2, 3 y 4"
												lErrorNumber = -1
											End If
										Case 38
											If (CInt(sPeriod) <> 1) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1"
												lErrorNumber = -1
											End If
										Case 39, 40
											If Not ((CInt(sPeriod) > 0) And (CInt(sPeriod) < 13)) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser del 1 al 12 (Enero a Diciembre)"
												lErrorNumber = -1
											End If
									End Select
								End If
							Case "PeriodYear"
								sPeriodYear = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If (Not IsEmpty(sPeriodYear)) And (Len(sPeriodYear) > 0) Then
									If Len(sPeriodYear) <> 4 Then
										sErrorDescription = "Introduzca el año con formato de 4 digitos"
										lErrorNumber = -1
									End If
								End If
							Case "Reasons"
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								aAbsenceComponent(S_REASONS_ABSENCE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ForJustificationID"
								If lErrorNumber = 0 Then
									sJustificationShortName = Right(("0000" & CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))), Len("0000"))
									If (Not IsEmpty(sJustificationShortName)) And (Len(sJustificationShortName) > 0) Then
										If VerifyExistenceOfRecordInDatabase(oADODBConnection, "Absences", "AbsenceID,StartDate,EndDate", CStr(N_NONE & "," & N_OPEN_MINIMUM & "," & N_OPEN_MAXIMUM), CStr(sJustificationShortName & "," & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & "," & "30000000"), oRecordset, sErrorDescription) Then
											lErrorNumber = -1
											lConceptError = lConceptError + 1
											sErrorDescription = "No existe identificador para la clave de la incidencia indicada: " & sJustificationShortName
										Else
											sQuery = "Select * from Absences Where (AbsenceShortName = '" & sJustificationShortName & "')"
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
											If Not oRecordset.EOF Then
												aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
											End If
										End If
									End If
								End If
						End Select
					Next
					If lErrorNumber = 0 Then
						Select Case lReasonID
							Case EMPLOYEES_EXTRAHOURS
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 10
							Case EMPLOYEES_SUNDAYS
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 16
								If Not IsSunday(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) Then
									sErrorDescription = "El concepto solo se puede registrar los domingos"
									lErrorNumber = -1
								End If
						End Select
						If (Len(sPeriodYear) > 0) AND (Len(sPeriod) > 0) Then
							aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = sPeriodYear & sPeriod
						Else
							Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
								Case 35, 37, 38
									sErrorDescription = "Para el registro de vacaciones debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
								Case 39
									sErrorDescription = "Para el registro del 'estimulo al trabajador del mes' debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
								Case 40
									sErrorDescription = "Para registrar 'sin derecho a estimulo por desempeño' debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
							End Select
							aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = 0
						End If
						If lErrorNumber = 0 Then
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo guardar la información del registro."
								If (lReasonID = 0) Or (lReasonID = 1) Then
									sAbsenceType = ""
									Call VerifyAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceType, sErrorDescription)
									If VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) And ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>21) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>22) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>23)) Then
										Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
											Case 41, 42, 43, 44, 45, 46, 47, 48, 49, 57, 58
												If (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = aAbsenceComponent(N_END_DATE_ABSENCE)) Or (aAbsenceComponent(N_END_DATE_ABSENCE) = 0) Then
													aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
												End If
											Case 50, 51, 54, 55, 56
												aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
											Case Else
												If aAbsenceComponent(N_END_DATE_ABSENCE) = 0 Then
													aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												End If
										End Select
										aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
										If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
											lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
										Else
											lErrorNumber = -1
										End If
									Else
										Select Case sAbsenceType
											Case "Justification"
												lErrorNumber = GetAbsenceAppliesToID(oRequest, oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
												aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
												If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
													If (VerifyExistenceOfAbsencesForJustification(oADODBConnection, aAbsenceComponent, aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), false, iAbsenceID, iActiveOriginal, sErrorDescription)) Then
														aAbsenceComponent(N_ACTIVE_ABSENCE) = 0
														lErrorNumber = AddJustification(oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
													Else
														lErrorNumber = -1
													End If
												Else
													lErrorNumber = -1
												End If
											Case Else
												aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												If ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=21) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=22) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=23)) Then
													If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) <> aAbsenceComponent(N_END_DATE_ABSENCE) Then
														lErrorNumber = -1
														sErrorDescription = "Para registrar vacaciones a empleados con tipo de jornada 2 se deben de registrar los días de manera individual."
													End If
												End If
												If lErrorNumber = 0 Then
													aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
													aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
													If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
														lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
													Else													
														lErrorNumber = -1
													End If
												End If
										End Select
									End If
								Else
									lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
								End If
							Else
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						Else
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesAbsencesFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesAdjustmentsFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         DocumentsForLicenses table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesAdjustmentsFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sMissingDateFormat
	Dim sPayrollDateFormat
	Dim sMissingDate
	Dim sPayrollDate
	Dim asInputDate
	Dim sFields
	Dim sQuery
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iDocumentDateYear
	Dim iDocumentDateMonth
	Dim sConceptShortName	
	Dim sError
	Dim iForPayrollIsActiveConstant

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "ConceptID"
						aiFieldsOrder(1) = iIndex & ",ConceptID"
					Case "ConceptAmount"
						aiFieldsOrder(2) = iIndex & ",ConceptAmount"
					Case "OcurredMissingDateYYYYMMDD"
						sMissingDateFormat = "YYYYMMDD"
						aiFieldsOrder(3) = iIndex & ",MissingDate"
					Case "OcurredMissingDateDDMMYYYY"
						sMissingDateFormat = "DDMMYYYY"
						aiFieldsOrder(3) = iIndex & ",MissingDate"
					Case "OcurredMissingDateMMDDYYYY"
						sMissingDateFormat= "MMDDYYYY"
						aiFieldsOrder(3) = iIndex & ",MissingDate"
					Case "OcurredPayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(4) = iIndex & ",PayrollDate"
					Case "OcurredPayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(4) = iIndex & ",PayrollDate"
					Case "OcurredMissingDateMMDDYYYY"
						sPayrollDateFormat= "MMDDYYYY"
						aiFieldsOrder(4) = iIndex & ",PayrollDate"
					Case "BeneficiaryName"
						aiFieldsOrder(5) = iIndex & ",BeneficiaryName"
				End Select
			End If
		Next
		sFields = "EmployeeID, ConceptID, ConceptAmount, MissingDate, PayrollDate, PaymentDate, ModifyDate, UserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de empleado."
		ElseIf InStr(1, sFields, "ConceptID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la clave del concepto."
		ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Cantidad a ajustar."
		ElseIf InStr(1, sFields, "MissingDate") > 0 Then 
			lErrorNumber = -1 
			sErrorDescription = "La información a registrar no contiene la Fecha de omisión de pago."
		ElseIf InStr(1, sFields, "PayrollDate") > 0 Then 
			lErrorNumber = -1 
			sErrorDescription = "La información a registrar no contiene la Fecha de aplicación de nómina."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					lErrorNumber = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ConceptAmount"
								If lErrorNumber = 0 Then
									sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <= 0 Then
										lErrorNumber = -1
										sErrorDescription = "El monto del reclamo no puede ser menor o igual a 0."
									End If
								End If
							Case "MissingDate"
								If lErrorNumber = 0 Then
									Select Case sMissingDateFormat
										Case "YYYYMMDD"
											sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
											sMissingDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
											sMissingDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
											sMissingDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sMissingDate) Then
										aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) = CLng(sMissingDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de omisión del pago en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(N_MISSING_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "PayrollDate"
								If lErrorNumber = 0 Then
									Select Case sPayrollDateFormat
										Case "YYYYMMDD"
											sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
											sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
											sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
											sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sPayrollDate) Then
										aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CLng(sPayrollDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la quincena de aplicación en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										Else
											If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
												iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
											ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
												iForPayrollIsActiveConstant = N_PAYROLL_FOR_FEATURES
											ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
												iForPayrollIsActiveConstant = 0
											End If
											If Not VerifyPayrollIsActive(oADODBConnection, aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), iForPayrollIsActiveConstant, sErrorDescription) Then
												lErrorNumber = -1
											End If
										End If
									End If
								End If
							Case "ConceptID"
								sConceptShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								If VerifyExistenceOfCatalogInDate(oADODBConnection, "Concepts", "ConceptShortName,StartDate,EndDate", CStr(sConceptShortName & "," & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & "," & "30000000"), "", oRecordset, sErrorDescription) Then
									aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptID").Value)
								Else
									lErrorNumber = -1
									sErrorDescription = "No se pudo obtener el concepto de la clave indicada."
								End If
							Case "BeneficiaryName"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
								aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
						End Select
					Next
					If lErrorNumber = 0 Then
						lErrorNumber = AddEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							sError = sErrorDescription
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sError & "<BR /><BR />"
						End If
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesAdjustmentsFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesAntiquitiesFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesConceptsLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesAntiquitiesFile"
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "ConceptAmount"
						aiFieldsOrder(1) = iIndex & ",ConceptAmount"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(2) = iIndex & ",StartDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(2) = iIndex & ",StartDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(2) = iIndex & ",StartDate"
				End Select
			End If
		Next

		sFields = "EmployeeID, ConceptAmount, StartDate, ConceptID, EndDate, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, StartUserID, EndUserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el monto a pagar."
		ElseIf InStr(1, sFields, "OcurredDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de inicio de aplicación."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sValues = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sFields, "ConceptID", "5"), "EndDate", "30000000"), "CurrencyID", "0"), "ConceptQttyID", "1"), "ConceptTypeID", "3"), "ConceptMinQttyID", "1"), "ConceptMin", "0"), "ConceptMaxQttyID", "1"), "ConceptMax", "0"), "AppliesToID", "-1"), "AbsenceTypeID", "1"), "ConceptOrder", "1"), "Active", "1"), "StartUserID", aLoginComponent(N_USER_ID_LOGIN)), "EndUserID", "-1")
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					sQuery = "Insert Into EmployeesConceptsLKP ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "OcurredDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case ""
							Case Else
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
						End Select
					Next
					sQuery = sQuery & sValues & ")"
					sErrorDescription = "No se pudo guardar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesAntiquitiesFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesAssignNumberFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         DocumentsForLicenses table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesAssignNumberFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormatDocument
	Dim sDateFormatStart
	Dim sDateFormatEnd
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iDocumentDateYear
	Dim iDocumentDateMonth
	Dim lEmployeeID

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeName"
						aiFieldsOrder(0) = iIndex & ",EmployeeName"
					Case "EmployeeLastName"
						aiFieldsOrder(1) = iIndex & ",EmployeeLastName"
					Case "EmployeeLastName2"
						aiFieldsOrder(2) = iIndex & ",EmployeeLastName2"
					Case "EmployeeTypeID"
						aiFieldsOrder(3) = iIndex & ",EmployeeTypeID"
					Case "RFC"
						aiFieldsOrder(4) = iIndex & ",RFC"
					Case "CURP"
						aiFieldsOrder(5) = iIndex & ",CURP"
					Case "EmployeeEmail"
						aiFieldsOrder(6) = iIndex & ",EmployeeEmail"
					Case "SocialSecurityNumber"
						aiFieldsOrder(7) = iIndex & ",SocialSecurityNumber"
					Case "CountryID"
						aiFieldsOrder(8) = iIndex & ",CountryID"
					Case "GenderID"
						aiFieldsOrder(9) = iIndex & ",GenderID"
					Case "MaritalStatusID"
						aiFieldsOrder(10) = iIndex & ",MaritalStatusID"
					Case "OcurredDocumentDateYYYYMMDD"
						sDateFormatDocument = "YYYYMMDD"
						aiFieldsOrder(11) = iIndex & ",BirthDate"
					Case "OcurredDocumentDateDDMMYYYY"
						sDateFormatDocument = "DDMMYYYY"
						aiFieldsOrder(11) = iIndex & ",BirthDate"
					Case "OcurredDocumentDateMMDDYYYY"
						sDateFormatDocument = "MMDDYYYY"
						aiFieldsOrder(11) = iIndex & ",BirthDate"
					Case "EmployeeAddress"
						aiFieldsOrder(12) = iIndex & ",EmployeeAddress"
					Case "EmployeeCity"
						aiFieldsOrder(13) = iIndex & ",EmployeeCity"
					Case "EmployeeZipCode"
						aiFieldsOrder(14) = iIndex & ",EmployeeZipCode"
					Case "StateID"
						aiFieldsOrder(15) = iIndex & ",StateID"
					Case "EmployeePhone"
						aiFieldsOrder(16) = iIndex & ",EmployeePhone"
					Case "OfficePhone"
						aiFieldsOrder(17) = iIndex & ",OfficePhone"
					Case "OfficeExt"
						aiFieldsOrder(18) = iIndex & ",OfficeExt"
					Case "DocumentNumber1"
						aiFieldsOrder(19) = iIndex & ",DocumentNumber1"
					Case "DocumentNumber2"
						aiFieldsOrder(20) = iIndex & ",DocumentNumber2"
					Case "DocumentNumber3"
						aiFieldsOrder(21) = iIndex & ",DocumentNumber3"
					Case "EmployeeActivityID"
						aiFieldsOrder(22) = iIndex & ",EmployeeActivityID"
				End Select
			End If
		Next

		sFields = ""
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nombre del empleado."
		ElseIf InStr(1, sFields, "EmployeeLastName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el apellido paterno del empleado."
		ElseIf InStr(1, sFields, "EmployeeTypeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la clave del tipo de tabulador."
		ElseIf InStr(1, sFields, "RFC") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el RFC."
		ElseIf InStr(1, sFields, "CURP") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el CURP."
		ElseIf InStr(1, sFields, "CountryID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el país."
		ElseIf InStr(1, sFields, "GenderID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el género."
		ElseIf InStr(1, sFields, "MaritalStatusID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el estado civil del empleado."
		ElseIf InStr(1, sFields, "EmployeeAddress") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el domicilio del empleado."
		ElseIf InStr(1, sFields, "EmployeeCity") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la ciudad del empleado."
		ElseIf InStr(1, sFields, "EmployeeZipCode") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el código postal del domicilio del empleado."
		ElseIf InStr(1, sFields, "StateID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el estado del domicilio del empleado."
		ElseIf InStr(1, sFields, "StateID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el estado del domicilio del empleado."
		ElseIf InStr(1, sFields, "BirthDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de nacimiento del empleado."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "BirthDate"
								Select Case sDateFormatDocument
									Case "YYYYMMDD"
										aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "EmployeeName"
								aEmployeeComponent(S_NAME_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeLastName"
								aEmployeeComponent(S_LAST_NAME_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeLastName2"
								aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeTypeID"
								aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "RFC"
								aEmployeeComponent(S_RFC_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "CURP"
								aEmployeeComponent(S_CURP_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeEmail"
								aEmployeeComponent(S_EMAIL_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "SocialSecurityNumber"
								aEmployeeComponent(S_SSN_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "CountryID"
								aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "GenderID"
								aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "MaritalStatusID"
								aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeAddress"
								aEmployeeComponent(S_ADDRESS_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeCity"
								aEmployeeComponent(S_CITY_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeZipCode"
								aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "StateID"
								aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeePhone"
								aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "OfficePhone"
								aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "OfficeExt"
								aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "DocumentNumber1"
								aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "DocumentNumber1"
								aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "DocumentNumber1"
								aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeActivityID"
								aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case ""
						End Select
					Next
					aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) = aEmployeeComponent(S_CURP_EMPLOYEE)
					aEmployeeComponent(S_PASSWORD_EMPLOYEE) = aEmployeeComponent(S_CURP_EMPLOYEE)
					sErrorDescription = "No se pudo obtener el número de empleado."
					Select Case CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE))
						Case 0,1,2,3,4
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=-1)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							aEmployeeComponent(N_ID_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
							aEmployeeComponent(S_NUMBER_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
						Case 6
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=5)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							aEmployeeComponent(N_ID_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
							aEmployeeComponent(S_NUMBER_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
						Case 5,7,8
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							aEmployeeComponent(N_ID_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
							aEmployeeComponent(S_NUMBER_EMPLOYEE)  = CLng(oRecordset.Fields("CurrentID").Value) + 1
					End Select
					If lErrorNumber = 0 Then
						sErrorDescription = "El RFC no es válido"
						If Len(aEmployeeComponent(S_RFC_EMPLOYEE))=13 Then
							sErrorDescription = "El CURP no es válido"
							If Len(aEmployeeComponent(S_CURP_EMPLOYEE))=18 Then
								sErrorDescription = "El RFC y el CURP son inconsistentes"
								If (InStr(1, Left(aEmployeeComponent(S_RFC_EMPLOYEE),10), Left(aEmployeeComponent(S_CURP_EMPLOYEE),10), vbBinaryCompare) > 0) Then
									sErrorDescription = "El RFC y la fecha de nacimiento son inconsistentes"
									If (InStr(1, Right(Left(aEmployeeComponent(S_RFC_EMPLOYEE),10),6), Right(CStr(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)),6), vbBinaryCompare) > 0) Then
										sErrorDescription = "No se pudo obtener el número para asignarlo al empleado."
										lEmployeeID = aEmployeeComponent(N_ID_EMPLOYEE)
										lErrorNumber = CheckExistencyOfEmployeeRFC(aEmployeeComponent, sErrorDescription)
										If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
											aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID
											aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
											aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 0
											lErrorNumber = AddEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
										Else
											lErrorNumber = L_ERR_DUPLICATED_RECORD
											sErrorDescription = "El empleado con número " & aEmployeeComponent(N_ID_EMPLOYEE) & " tiene el mismo RFC, CURP o el mismo nombre completo."
										End If
									Else
										lErrorNumber = N_ERROR_LEVEL
									End If
								Else
									lErrorNumber = N_ERROR_LEVEL
								End If
							Else
								lErrorNumber = N_ERROR_LEVEL
							End If
						Else
							lErrorNumber = N_ERROR_LEVEL
						End If
						If lErrorNumber <> 0 Then
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex + 1) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesAssignNumberFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesBankAccountFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         BankAccounts table.
'Inputs:  oADODBConnection, lReasonID, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesBankAccountFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sEndDateFormat
	Dim sPayrollDateFormat
	Dim asInputDate
	Dim sFields
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim sPayrollDate
	Dim sEndDate
	Dim sBankShortName

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeID"
					Case "BankID"
						aiFieldsOrder(iIndex) = iIndex & ",BankID"
					Case "AccountNumber"
						aiFieldsOrder(iIndex) = iIndex & ",AccountNumber"
					Case "EndDateYYYYMMDD"
						sEndDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateDDMMYYYY"
						sEndDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateMMDDYYYY"
						sEndDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
				End Select
			End If
		Next
		sFields = "EmployeeID, BankID, AccountNumber, PayrollDate, "
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then
				sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
			End If
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "BankID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el id del banco."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de termino."
		ElseIf InStr(1, sFields, "AccountNumber") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de cuenta."
		ElseIf InStr(1, sFields, "PayrollDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la quincena de aplicación."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = -1
				sEndDate = Empty
				If Len(asFileContents(iIndex)) > 0 Then
					lErrorNumber = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sErrorDescription = "No existe empleado registrado con el número indicado."
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									sErrorDescription = "No existe empleado registrado con el número indicado."
								End If
							Case "BankID"
								If lErrorNumber = 0 Then
									sBankShortName = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									lErrorNumber = CheckExistencyOfBankShortName(aEmployeeComponent, sBankShortName, sErrorDescription)
								End If
							Case "PayrollDate"
								If lErrorNumber = 0 Then
									Select Case sPayrollDateFormat
										Case "YYYYMMDD"
											sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sPayrollDate) Then
										aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CLng(sPayrollDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
										lErrorNumber = -1
									Else
										aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										Else
											If Not VerifyPayrollIsActive(oADODBConnection, aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), N_PAYROLL_FOR_BANK, sErrorDescription) Then
												lErrorNumber = -1
											End If
										End If
									End If
								End If
							Case "EndDate"
								If lErrorNumber = 0 Then
									Select Case sEndDateFormat
										Case "YYYYMMDD"
											sEndDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sEndDate) Then
										aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(sEndDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "AccountNumber"
								If lErrorNumber = 0 Then
									If Len(Cstr(asFileRow(Cstr(aiFieldsOrder(jIndex)(0))))) > 0 Then
										aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = CStr(asFileRow(Cstr(aiFieldsOrder(jIndex)(0))))
									Else
										sErrorDescription = "El número de cuenta no puede estar vacio."
										lErrorNumber = -1
									End If
								End If
						End Select
					Next
					If lErrorNumber = 0 Then
						If (aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0) Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
						If (aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) > aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) Then
							sErrorDescription = "La fecha de inicio no puede ser mayor a la fecha de fin."
							lErrorNumber = -1
						End If
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información del registro."
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
							lErrorNumber = AddEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesBankAccountFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesBeneficiariesDebitFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesConceptsLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesBeneficiariesDebitFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sStartDateFormat
	Dim sEndDateFormat
	Dim sPayrollDateFormat
	Dim asInputDate
	Dim sFields
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iStartDate
	Dim sStartDate
	Dim iEndDate
	Dim lEndDate
	Dim sEndDate
	Dim iPayrollDate
	Dim sPayrollDate

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "StartDateYYYYMMDD"
						sStartDateFormat = "YYYYMMDD"
						aiFieldsOrder(1) = iIndex & ",StartDate"
					Case "StartDateDDMMYYYY"
						sStartDateFormat = "DDMMYYYY"
						aiFieldsOrder(1) = iIndex & ",StartDate"
					Case "StartDateMMDDYYYY"
						sStartDateFormat= "MMDDYYYY"
						aiFieldsOrder(1) = iIndex & ",StartDate"
					Case "EndDateYYYYMMDD"
						sEndDateFormat = "YYYYMMDD"
						aiFieldsOrder(2) = iIndex & ",EndDate"
					Case "EndDateDDMMYYYY"
						sEndDateFormat = "DDMMYYYY"
						aiFieldsOrder(2) = iIndex & ",EndDate"
					Case "EndDateMMDDYYYY"
						sEndDateFormat= "MMDDYYYY"
						aiFieldsOrder(2) = iIndex & ",EndDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat= "MMDDYYYY"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "ConceptAmount"
						aiFieldsOrder(4) = iIndex & ",ConceptAmount"
					Case "BeneficiaryNumber"
						aiFieldsOrder(5) = iIndex & ",BeneficiaryNumber"
					Case "ConceptComments"
						aiFieldsOrder(6) = iIndex & ",ConceptComments"
				End Select
			End If
		Next
		sFields = "EmployeeID, StartDate, EndDate, PayrollDate, ConceptAmount, BeneficiaryNumber, "
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de empleado."
		ElseIf InStr(1, sFields, "StartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Fecha de inicio."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Fecha de termino."
		ElseIf InStr(1, sFields, "PayrollDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha aplicación en nómina."
		ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el importe."
		ElseIf InStr(1, sFields, "BeneficiaryNumber") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de beneficiario."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					lErrorNumber = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								sErrorDescription = "No existe el empleado indicado"
								If lErrorNumber = 0 Then
									If Not VerifyExistenceOfEmployeesBeneficiary(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
										sErrorDescription = "El empleado no tiene registrados beneficiarios para capturar este concepto"
										lErrorNumber = -1
									End If
								End If
							Case "StartDate"
								Select Case sStartDateFormat
									Case "YYYYMMDD"
										iStartDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))										
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "EndDate"
								Select Case sEndDateFormat
									Case "YYYYMMDD"
										sEndDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))										
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
								If Not IsEmpty(sEndDate) Then
									iEndDate = CLng(sEndDate)
								Else
									If (Err.Number <> 0) Then
										Err.Clear
									End If
								End If
								If (Err.Number <> 0) Then
									Err.Clear
									sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
									lErrorNumber = -1
								Else
									If iEndDate = 0 Then iEndDate = 30000000
									If lErrorNumber = 0 Then
										If CLng(iEndDate) < CLng(iStartDate) Then
											sErrorDescription = "La fecha de fin no puede ser menor a la fecha de inicio"
											lErrorNumber = -1
										End If
									End If
								End If
							Case "PayrollDate"
								Select Case sPayrollDateFormat
									Case "YYYYMMDD"
										sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
								If Not IsEmpty(sPayrollDate) Then
									iPayrollDate = CLng(sPayrollDate)
								Else
									If (Err.Number <> 0) Then
										Err.Clear
									End If
								End If
								If (Err.Number <> 0) Then
									Err.Clear
									sErrorDescription = "Introduzca la quincena de aplicación en un formato correcto."
									lErrorNumber = -1
								Else
									If lErrorNumber = 0 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where PayrollDate = '" & iPayrollDate & "' And (IsClosed<>1)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If oRecordset.EOF Then
												lErrorNumber = -1
												sErrorDescription = "La fecha de nómina no esta registrada o no esta abierta"
											End If
										End If
									End If
								End If
							Case "ConceptAmount"
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) =  asFileRow(CDbl(aiFieldsOrder(jIndex)(0)))
							Case "BeneficiaryNumber"
								aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
						End Select
					Next
					If lErrorNumber = 0 Then					
						aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = iStartDate
						aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = iPayrollDate
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sFileName
						aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							sErrorDescription = "No se pudo guardar la información del registro de adeudo de pensión alimenticia"
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If	
				End If
			Next		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesBeneficiariesDebitFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesChildrenFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesChildrenLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesChildrenFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "ChildName"
						aiFieldsOrder(1) = iIndex & ",ChildName"
					Case "ChildLastName"
						aiFieldsOrder(2) = iIndex & ",ChildLastName"
					Case "ChildLastName2"
						aiFieldsOrder(3) = iIndex & ",ChildLastName2"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(4) = iIndex & ",ChildBirthDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(4) = iIndex & ",ChildBirthDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(4) = iIndex & ",ChildBirthDate"
				End Select
			End If
		Next

		sFields = "EmployeeID, ChildName, ChildLastName, ChildBirthDate, ChildLastName2, ChildID, ChildEndDate, SchoolarshipID, RegistrationDate, UserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "ChildName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nombre del hijo(a)."
		ElseIf InStr(1, sFields, "ChildBirthDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de nacimiento de los hijos."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) = "''"
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "ChildBirthDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "ChildName"
								aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ChildLastName"
								aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ChildLastName2"
								aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case Else
						End Select
					Next
					lErrorNumber = SaveEmployeeChildren(aEmployeeComponent, "ChildrenSchoolarships", sErrorDescription)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesChildrenFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesExtraHoursFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesAbsencesLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesExtraHoursFile"
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "AbsenceHours"
						aiFieldsOrder(1) = iIndex & ",AbsenceHours"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(2) = iIndex & ",OcurredDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(2) = iIndex & ",OcurredDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(2) = iIndex & ",OcurredDate"
					Case "DocumentNumber"
						aiFieldsOrder(3) = iIndex & ",DocumentNumber"
					Case "Reasons"
						aiFieldsOrder(4) = iIndex & ",Reasons"
				End Select
			End If
		Next

		sFields = "EmployeeID, AbsenceHours, OcurredDate, EndDate, AbsenceID, RegistrationDate, DocumentNumber, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then
				sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
			ElseIf StrComp(aiFieldsOrder(iIndex)(1), "OcurredDate", vbBinaryCompare) = 0 Then
				sFields = Replace(sFields, "EndDate, ", "")
			End If
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "AbsenceHours") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de horas."
		ElseIf InStr(1, sFields, "OcurredDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las fechas de las horas extras."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sValues = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sFields, "RegistrationDate", sDate), "AbsenceID", "201"), "JustificationID", "-1"), "AppliesForPunctuality", "1"), "AddUserID", aLoginComponent(N_USER_ID_LOGIN)), "AppliedDate", "0"), "Removed,", "0,"), "RemoveUserID", "-1"), "RemovedDate", "0"), "AppliedRemoveDate", "0"), "EndDate,", "")
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)

					sQuery = "Insert Into EmployeesAbsencesLKP ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "OcurredDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", " & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", " & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", " & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case "DocumentNumber", "Reasons"
								sQuery = sQuery & "'" & Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´") & "', "
							Case ""
							Case Else
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
						End Select
					Next
					If InStr(1, sValues, "DocumentNumber") > 0 Then
						sValues = Replace(sValues, "DocumentNumber", "'S/N'")
					End If
					If InStr(1, sValues, "Reasons") > 0 Then
						sValues = Replace(sValues, "Reasons", "''")
					End If
					sQuery = sQuery & sValues & ")"
					sErrorDescription = "No se pudo guardar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesExtraHoursFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesSafeSeparationFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesConceptsLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesSafeSeparationFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sStartDateFormat
	Dim sPayrollDateFormat
	Dim asInputDate
	Dim sFields
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iStartDate
	Dim iPayrollDate	

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "ConceptID"
						aiFieldsOrder(1) = iIndex & ",ConceptID"
					Case "OcurredStartDateYYYYMMDD"
						sStartDateFormat = "YYYYMMDD"
						aiFieldsOrder(2) = iIndex & ",StartDate"
					Case "OcurredStartDateDDMMYYYY"
						sStartDateFormat = "DDMMYYYY"
						aiFieldsOrder(2) = iIndex & ",StartDate"
					Case "OcurredStartDateMMDDYYYY"
						sStartDateFormat= "MMDDYYYY"
						aiFieldsOrder(2) = iIndex & ",StartDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat= "MMDDYYYY"
						aiFieldsOrder(3) = iIndex & ",PayrollDate"
					Case "ConceptAmount"
						aiFieldsOrder(4) = iIndex & ",ConceptAmount"
					Case "ConceptQttyID"
						aiFieldsOrder(5) = iIndex & ",ConceptQttyID"
				End Select
			End If
		Next
		sFields = "EmployeeID, StartDate, PayrollDate, ConceptAmount, ConceptQttyID, "
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de empleado."
		ElseIf InStr(1, sFields, "ConceptID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de Concepto."
		ElseIf InStr(1, sFields, "StartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la Fecha de inicio."
		ElseIf InStr(1, sFields, "PayrollDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de nómina."
		ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
			lErrorNumber = -1
			If sAction = "EmployeesSafeSeparation" Then
				sErrorDescription = "La información a registrar no contiene el porcentaje."
			Else
				sErrorDescription = "La información a registrar no contiene la cantidad ($ o %)."
			End If
		ElseIf InStr(1, sFields, "ConceptQttyID") > 0 Then
			If lReasonID <> EMPLOYEES_SAFE_SEPARATION Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene el tipo de cantidad ($ o %)."
			End If
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sErrorDescription = "No existe el empleado indicado"
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							Case "StartDate"								
								Select Case sStartDateFormat
									Case "YYYYMMDD"
										iStartDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))										
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "PayrollDate"
								Select Case sPayrollDateFormat
									Case "YYYYMMDD"
										iPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										iPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "ConceptAmount"
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) =  asFileRow(CLng(aiFieldsOrder(jIndex)(0)))
							Case "ConceptQttyID"
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
						End Select
					Next
					If lReasonID = EMPLOYEES_SAFE_SEPARATION Then
						If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 4 And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 5) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 10) Then
							lErrorNumber = -1
							sErrorDescription = "El porcentaje para SI solo puede ser 2, 4, 5 y 10."
						End If 
					Else
						sQuery = "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=120) And (StartDate>" & aEmployeeComponent(iStartDate) & ") And (EndDate=30000000)"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If ((lErrorNumber <> 0) And (Not oRecordset.EOF)) Then
							sErrorDescription = "Para capturar el seguro adicional, debe de estar registrado el concepto SI."
							lErrorNumber = -1
						ElseIf (aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 100) Then
							lErrorNumber = -1
							sErrorDescription = "El porcentaje para AE no puede ser mayor al 100 %."
						End If
					End If
					If lErrorNumber = 0 Then					
						aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = iStartDate
						aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = iPayrollDate
						aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,3"
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sFileName
						aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
						If lReasonID = EMPLOYEES_SAFE_SEPARATION Then
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120
							sMessage = "seguro de separación individualizado."
						Else
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87
							sMessage = "seguro adicional de separación individualizado."
						End If
						lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							sErrorDescription = "No se pudo guardar la información del registro por " & sMessage
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If	
				End If
			Next		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesSafeSeparationFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesSpecialJourneysFile(oADODBConnection, lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesAbsencesLKP table.
'Inputs:  oADODBConnection, lReasonID, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesSpecialJourneysFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim sEndDateFormat
	Dim sPayrollDateFormat
	Dim sStartDate
	Dim sEndDate
	Dim sPayrollDate
	Dim asInputDate
	Dim sFields
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iEmployeeTypeID
	Dim iPositionTypeID
	Dim sPeriod
	Dim sPeriodYear
	Dim sAbsenceIDs
	Dim sAbsenceType
	Dim iAbsenceID
	Dim iActiveOriginal
	Dim sAbsenceShortName
	Dim sJustificationShortName

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeID"
					Case "AbsenceID"
						aiFieldsOrder(iIndex) = iIndex & ",AbsenceID"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",OcurredDate"
					Case "EndDateYYYYMMDD"
						sEndDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateDDMMYYYY"
						sEndDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateMMDDYYYY"
						sEndDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "VacationPeriod"
						aiFieldsOrder(iIndex) = iIndex & ",VacationPeriod"
					Case "PeriodYear"
						aiFieldsOrder(iIndex) = iIndex & ",PeriodYear"
					Case "Reasons"
						aiFieldsOrder(iIndex) = iIndex & ",Reasons"
					Case "DocumentNumber"
						aiFieldsOrder(iIndex) = iIndex & ",DocumentNumber"
					Case "AbsenceHours"
						aiFieldsOrder(iIndex) = iIndex & ",AbsenceHours"
					Case "JustificationID"
						aiFieldsOrder(iIndex) = iIndex & ",JustificationID"
					Case "AppliesForPunctuality"
						aiFieldsOrder(iIndex) = iIndex & ",AppliesForPunctuality"
					Case "ForJustificationID"
						aiFieldsOrder(iIndex) = iIndex & ",ForJustificationID"
				End Select
			End If
		Next
		Select Case lReasonID
			Case EMPLOYEES_EXTRAHOURS
				sFields = "EmployeeID, OcurredDate, PayrollDate, AbsenceHours, Reasons"
			Case EMPLOYEES_SUNDAYS
				sFields = "EmployeeID, OcurredDate, PayrollDate, Reasons"
			Case 0, 1
				sFields = "EmployeeID, AbsenceID, OcurredDate, PayrollDate"
			Case Else
				sFields = "EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate"
		End Select
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "AbsenceID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el tipo de incidencias."
		ElseIf InStr(1, sFields, "OcurredDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de las incidencias."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de termino de las incidencias."
		ElseIf (InStr(1, sFields, "AbsenceHours") > 0) And (lReasonID = EMPLOYEES_EXTRAHOURS) Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las horas extras."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					sEndDate = Empty
					sPayrollDate = Empty
					lErrorNumber = 0
					aAbsenceComponent(N_END_DATE_ABSENCE) = 0
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								aEmployeeComponent(N_ID_EMPLOYEE) = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								sErrorDescription = "No existe el empleado indicado"
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									sErrorDescription = "Error al verificar la existencia del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID = '" & aEmployeeComponent(N_ID_EMPLOYEE) & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									iEmployeeTypeID = oRecordset.Fields("EmployeeTypeID").Value
									iPositionTypeID = oRecordset.Fields("PositionTypeID").Value
									oRecordset.Close
								End If
							Case "AbsenceID"
								If lErrorNumber = 0 Then
									sAbsenceShortName = Right(("0000" & CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))), Len("0000"))
									sQuery = "Select * from Absences Where (AbsenceShortName = '" & sAbsenceShortName & "')"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If oRecordset.EOF Then
											lErrorNumber = -1
											lConceptError = lConceptError + 1
											sErrorDescription = "No existe identificador para la clave de la incidencia indicada: " & sAbsenceShortName
										Else
											aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("AbsenceID").Value)
											aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
										End If
										oRecordset.Close
									End If
								End If
							Case "OcurredDate"
								If lErrorNumber = 0 Then
									Select Case sDateFormat
										Case "YYYYMMDD"
											sStartDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(sStartDate)
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "EndDate"
								If lErrorNumber = 0 Then
									Select Case sEndDateFormat
										Case "YYYYMMDD"
											sEndDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sEndDate) Then
										aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(sEndDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_END_DATE_ABSENCE) = aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), sErrorDescription) Then
											lErrorNumber = -1
										End If
									End If
								End If
							Case "PayrollDate"
								If lErrorNumber = 0 Then
									Select Case sPayrollDateFormat
										Case "YYYYMMDD"
											sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sPayrollDate) Then
										aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CLng(sPayrollDate)
									Else
										If (Err.Number <> 0) Then
											Err.Clear
										End If
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la quincena de aplicación en un formato correcto."
										lErrorNumber = -1
									Else
										aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
										If Not VerifyIfUploadMonthDateIsCorrect(aAbsenceComponent(N_APPLIED_DATE_ABSENCE), sErrorDescription) Then
											lErrorNumber = -1
										Else
											If (lReasonID = 0) Or (lReasonID = 1) Then
												If Not VerifyPayrollIsActive(oADODBConnection, aAbsenceComponent(N_APPLIED_DATE_ABSENCE), N_PAYROLL_FOR_ABSENCES, sErrorDescription) Then
													lErrorNumber = -1
												End If
											Else
												If Not VerifyPayrollIsActive(oADODBConnection, aAbsenceComponent(N_APPLIED_DATE_ABSENCE), N_PAYROLL_FOR_FEATURES, sErrorDescription) Then
													lErrorNumber = -1
												End If
											End If
										End If
									End If
								End If
							Case "DocumentNumber"
								aEmployeeComponent(xxxx) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "AbsenceHours"
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								Select Case lReasonID
									Case EMPLOYEES_EXTRAHOURS
										If lErrorNumber = 0 Then
											If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 1) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 3) Then
												sErrorDescription = "Solamente se pueden registrar 1, 2 o 3 horas extras en un día"
												lErrorNumber = -1
											End If
										End If
									Case EMPLOYEES_SUNDAYS
								End Select
							Case "JustificationID"
								aEmployeeComponent(N_CONCEPT_JUSTIFICATION_ID_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "AppliesForPunctuality"
								aEmployeeComponent(N_CONCEPT_FOR_PUNCTUALITY_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "VacationPeriod"
								sPeriod = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If (Not IsEmpty(sPeriod)) And (Len(sPeriod) > 0) Then
									Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
										Case 35
											If (CInt(sPeriod) <> 1) And (CInt(sPeriod) <> 2) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1 y 2"
												lErrorNumber = -1
											End If
										Case 37
											If (CInt(sPeriod) <> 1) And (CInt(sPeriod) <> 2) And (CInt(sPeriod) <> 3) And (CInt(sPeriod) <> 4) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1, 2, 3 y 4"
												lErrorNumber = -1
											End If
										Case 38
											If (CInt(sPeriod) <> 1) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser 1"
												lErrorNumber = -1
											End If
										Case 39, 40
											If Not ((CInt(sPeriod) > 0) And (CInt(sPeriod) < 13)) Then
												sErrorDescription = "Los periodos para este tipo de incidencia solamente pueden ser del 1 al 12 (Enero a Diciembre)"
												lErrorNumber = -1
											End If
									End Select
								End If
							Case "PeriodYear"
								sPeriodYear = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If (Not IsEmpty(sPeriodYear)) And (Len(sPeriodYear) > 0) Then
									If Len(sPeriodYear) <> 4 Then
										sErrorDescription = "Introduzca el año con formato de 4 digitos"
										lErrorNumber = -1
									End If
								End If
							Case "Reasons"
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								aAbsenceComponent(S_REASONS_ABSENCE) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ForJustificationID"
								If lErrorNumber = 0 Then
									sJustificationShortName = Right(("0000" & CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))), Len("0000"))
									If (Not IsEmpty(sJustificationShortName)) And (Len(sJustificationShortName) > 0) Then
										If VerifyExistenceOfRecordInDatabase(oADODBConnection, "Absences", "AbsenceID,StartDate,EndDate", CStr(N_NONE & "," & N_OPEN_MINIMUM & "," & N_OPEN_MAXIMUM), CStr(sJustificationShortName & "," & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & "," & "30000000"), oRecordset, sErrorDescription) Then
											lErrorNumber = -1
											lConceptError = lConceptError + 1
											sErrorDescription = "No existe identificador para la clave de la incidencia indicada: " & sJustificationShortName
										Else
											sQuery = "Select * from Absences Where (AbsenceShortName = '" & sJustificationShortName & "')"
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
											If Not oRecordset.EOF Then
												aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
											End If
										End If
									End If
								End If
						End Select
					Next
					If lErrorNumber = 0 Then
						Select Case lReasonID
							Case EMPLOYEES_EXTRAHOURS
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 10
							Case EMPLOYEES_SUNDAYS
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 16
								If Not IsSunday(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) Then
									sErrorDescription = "El concepto solo se puede registrar los domingos"
									lErrorNumber = -1
								End If
						End Select
						If (Len(sPeriodYear) > 0) AND (Len(sPeriod) > 0) Then
							aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = sPeriodYear & sPeriod
						Else
							Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
								Case 35, 37, 38
									sErrorDescription = "Para el registro de vacaciones debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
								Case 39
									sErrorDescription = "Para el registro del 'estimulo al trabajador del mes' debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
								Case 40
									sErrorDescription = "Para registrar 'sin derecho a estimulo por desempeño' debe indicar el año y el periodo en el que aplican."
									lErrorNumber = -1
							End Select
							aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = 0
						End If
						If lErrorNumber = 0 Then
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo guardar la información del registro."
								If (lReasonID = 0) Or (lReasonID = 1) Then
									sAbsenceType = ""
									Call VerifyAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceType, sErrorDescription)
									If VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) And ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>21) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>22) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>23)) Then
										Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
											Case 41, 42, 43, 44, 45, 46, 47, 48, 49, 57, 58
												If (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = aAbsenceComponent(N_END_DATE_ABSENCE)) Or (aAbsenceComponent(N_END_DATE_ABSENCE) = 0) Then
													aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
												End If
											Case 50, 51, 54, 55, 56
												aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
											Case Else
												If aAbsenceComponent(N_END_DATE_ABSENCE) = 0 Then
													aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												End If
										End Select
										aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
										If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
											lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
										Else
											lErrorNumber = -1
										End If
									Else
										Select Case sAbsenceType
											Case "Justification"
												lErrorNumber = GetAbsenceAppliesToID(oRequest, oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
												aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
												If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
													If (VerifyExistenceOfAbsencesForJustification(oADODBConnection, aAbsenceComponent, aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), false, iAbsenceID, iActiveOriginal, sErrorDescription)) Then
														aAbsenceComponent(N_ACTIVE_ABSENCE) = 0
														lErrorNumber = AddJustification(oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
													Else
														lErrorNumber = -1
													End If
												Else
													lErrorNumber = -1
												End If
											Case Else
												aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
												If ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=21) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=22) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=23)) Then
													If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) <> aAbsenceComponent(N_END_DATE_ABSENCE) Then
														lErrorNumber = -1
														sErrorDescription = "Para registrar vacaciones a empleados con tipo de jornada 2 se deben de registrar los días de manera individual."
													End If
												End If
												If lErrorNumber = 0 Then
													aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
													aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
													If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
														lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
													Else													
														lErrorNumber = -1
													End If
												End If
										End Select
									End If
								Else
									lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
								End If
							Else
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						Else
							sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesSpecialJourneysFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesSundaysFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesAbsencesLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesSundaysFile"
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "EmployeeID"
						aiFieldsOrder(0) = iIndex & ",EmployeeID"
					Case "OcurredDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(1) = iIndex & ",OcurredDate"
					Case "OcurredDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(1) = iIndex & ",OcurredDate"
					Case "OcurredDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(1) = iIndex & ",OcurredDate"
					Case "DocumentNumber"
						aiFieldsOrder(2) = iIndex & ",DocumentNumber"
					Case "Reasons"
						aiFieldsOrder(3) = iIndex & ",Reasons"
				End Select
			End If
		Next

		sFields = "EmployeeID, OcurredDate, EndDate, AbsenceID, RegistrationDate, DocumentNumber, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, AbsenceHours"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then
				sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
			ElseIf StrComp(aiFieldsOrder(iIndex)(1), "OcurredDate", vbBinaryCompare) = 0 Then
				sFields = Replace(sFields, "EndDate, ", "")
			End If
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "OcurredDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las fechas de los domingos."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sValues = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sFields, "RegistrationDate", sDate), "AbsenceID", "202"), "JustificationID", "-1"), "AppliesForPunctuality", "1"), "AddUserID", aLoginComponent(N_USER_ID_LOGIN)), "AppliedDate", "0"), "Removed,", "0,"), "RemoveUserID", "-1"), "RemovedDate", "0"), "AppliedRemoveDate", "0"), "EndDate,", ""), "AbsenceHours", "0")
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)

					sQuery = "Insert Into EmployeesAbsencesLKP ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "OcurredDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", " & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", " & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", " & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case "DocumentNumber", "Reasons"
								sQuery = sQuery & "'" & Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´") & "', "
							Case ""
							Case Else
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
						End Select
					Next
					If InStr(1, sValues, "DocumentNumber") > 0 Then
						sValues = Replace(sValues, "DocumentNumber", "'S/N'")
					End If
					If InStr(1, sValues, "Reasons") > 0 Then
						sValues = Replace(sValues, "Reasons", "''")
					End If
					sQuery = sQuery & sValues & ")"
					sErrorDescription = "No se pudo guardar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadEmployeesSundaysFile = lErrorNumber
	Err.Clear
End Function

Function UploadJobsFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         Employees, EmployeesHistoryList, Jobs tables.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadJobsFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sStartDateFormat
	Dim sEndDateFormat
	Dim asInputDate
	Dim sFields
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sQuery
	Dim lEndDate
	Dim lStartDate
	Dim lJobID
	Dim lServiceID
	Dim sServiceShortName
	Dim lShiftID
	Dim sShiftShortName
	Dim lJourneyID
	Dim sJourneyShortName
	Dim lAreaID
	Dim sAreaShortName
	Dim lPaymentCenterID
	Dim sPaymentCenterShortName
	Dim lJobTypeID
	Dim sJobTypeShortName
	Dim lPositionID
	Dim sPositionShotName
	Dim lLevelID
	Dim sLevelShortName
	Dim lGroupGradeLevelID
	Dim sGroupGradeLevelShortName
	Dim lIntegrationID
	Dim lClassificationID
	Dim sErrorUpload
	Dim sErrorQueries
	Dim lStatusID
	Dim sStatusShortName
	Dim bAddConsecutive

	sErrorUpload = ""
	lAreaID = 0
	lJobTypeID = 0
	lJourneyID = 0
	lPaymentCenterID = 0
	lPositionID = 0
	lServiceID = 0
	lShiftID = 0
	lStartDate = 0
	lEndDate = 0
	lLevelID = 0
	lGroupGradeLevelID = 0
	lIntegrationID = 0
	lClassificationID = 0
	lStatusID = -1
	sStatusShortName = ""
	
	
	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 2)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case lReasonID
					Case 54
						Select Case oRequest("Column" & (iIndex + 1)).Item
						    Case "JobID"
						            aiFieldsOrder(0) = iIndex & ",JobID"
						    Case "ServiceID"
						            aiFieldsOrder(1) = iIndex & ",ServiceID"
						End Select
					Case 60
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "JobID"
								aiFieldsOrder(0) = iIndex & ",JobID"
							Case "AreaID"
								aiFieldsOrder(1) = iIndex & ",AreaID"
							Case "PaymentCenterID"
								aiFieldsOrder(2) = iIndex & ",PaymentCenterID"
							Case "ServiceID"
								aiFieldsOrder(3) = iIndex & ",ServiceID"
							Case "JourneyID"
								aiFieldsOrder(4) = iIndex & ",JourneyID"
							Case "ShiftID"
								aiFieldsOrder(5) = iIndex & ",ShiftID"
							Case "StatusID"
								aiFieldsOrder(6) = iIndex & ",StatusID"
							Case "OcurredStartDateYYYYMMDD"
								sStartDateFormat = "YYYYMMDD"
								aiFieldsOrder(7) = iIndex & ",StartDate"
							Case "OcurredStartDateDDMMYYYY"
								sStartDateFormat = "DDMMYYYY"
								aiFieldsOrder(7) = iIndex & ",StartDate"
							Case "OcurredStartDateMMDDYYYY"
								sStartDateFormat = "MMDDYYYY"
								aiFieldsOrder(7) = iIndex & ",StartDate"
						End Select
					Case 61
						Select Case oRequest("Column" & (iIndex + 1)).Item
						    Case "JobID"
					            aiFieldsOrder(0) = iIndex & ",JobID"
							Case "PositionID"
								aiFieldsOrder(1) = iIndex & ",PositionID"
						    Case "LevelID"
								aiFieldsOrder(2) = iIndex & ",LevelID"
							Case "GroupGradeLevelID"
								aiFieldsOrder(3) = iIndex & ",GroupGradeLevelID"
							Case "IntegrationID"
								aiFieldsOrder(4) = iIndex & ",IntegrationID"
							Case "ClassificationID"
								aiFieldsOrder(5) = iIndex & ",ClassificationID"
						    Case "OcurredStartDateYYYYMMDD"
						            sStartDateFormat = "YYYYMMDD"
						            aiFieldsOrder(6) = iIndex & ",StartDate"
						    Case "OcurredStartDateDDMMYYYY"
						            sStartDateFormat = "DDMMYYYY"
						            aiFieldsOrder(6) = iIndex & ",StartDate"
						    Case "OcurredStartDateMMDDYYYY"
						            sStartDateFormat = "MMDDYYYY"
						            aiFieldsOrder(6) = iIndex & ",StartDate"
						End Select
					Case Else
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "PositionID"
								aiFieldsOrder(0) = iIndex & ",PositionID"
						    Case "LevelID"
								aiFieldsOrder(1) = iIndex & ",LevelID"
							Case "GroupGradeLevelID"
								aiFieldsOrder(2) = iIndex & ",GroupGradeLevelID"
							Case "IntegrationID"
								aiFieldsOrder(3) = iIndex & ",IntegrationID"
							Case "ClassificationID"
								aiFieldsOrder(4) = iIndex & ",ClassificationID"
							Case "ServiceID"
						            aiFieldsOrder(5) = iIndex & ",ServiceID"
							Case "JourneyID"
						            aiFieldsOrder(6) = iIndex & ",JourneyID"
						    Case "ShiftID"
						            aiFieldsOrder(7) = iIndex & ",ShiftID"    
						    Case "JobTypeID"
						            aiFieldsOrder(8) = iIndex & ",JobTypeID"
							Case "AreaID"
						            aiFieldsOrder(9) = iIndex & ",AreaID"
						    Case "PaymentCenterID"
						            aiFieldsOrder(10) = iIndex & ",PaymentCenterID"
						    Case "JobID"
						            aiFieldsOrder(11) = iIndex & ",JobID"
						    Case "OcurredStartDateYYYYMMDD"
						            sStartDateFormat = "YYYYMMDD"
						            aiFieldsOrder(12) = iIndex & ",StartDate"
						    Case "OcurredStartDateDDMMYYYY"
						            sStartDateFormat = "DDMMYYYY"
						            aiFieldsOrder(12) = iIndex & ",StartDate"
						    Case "OcurredStartDateMMDDYYYY"
						            sStartDateFormat = "MMDDYYYY"
						            aiFieldsOrder(12) = iIndex & ",StartDate"
						    Case "OcurredEndDateYYYYMMDD"
						            sEndDateFormat = "YYYYMMDD"
						            aiFieldsOrder(13) = iIndex & ",EndDate"
						    Case "OcurredEndDateDDMMYYYY"
						            sEndDateFormat = "DDMMYYYY"
						            aiFieldsOrder(13) = iIndex & ",EndDate"
						    Case "OcurredEndDateMMDDYYYY"
						            sEndDateFormat = "MMDDYYYY"
						            aiFieldsOrder(13) = iIndex & ",EndDate"
						End Select
					End Select
			End If
		Next

		sFields = ", "
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de empleado."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				ReDim Preserve aJobComponent(N_JOB_COMPONENT_SIZE)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sErrorUpload = ""
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "AreaID"
								lAreaID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sAreaShortName = Right(("00000" & lAreaID), Len("00000"))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID From Areas Where AreaCode = '" & sAreaShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lAreaID = CLng(oRecordset.Fields("AreaID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del centro de trabajo no existe.<BR />"
									End If
								End If
							Case "ClassificationID"
								If Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) > 0 Then
									lClassificationID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								Else
									lClassificationID = -1
								End If
							Case "GroupGradeLevelID"
								If Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) > 0 Then
									lGroupGradeLevelID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									sGroupGradeLevelShortName = Right("000" & lGroupGradeLevelID, 3)
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select GroupGradeLevelID From GroupGradeLevels Where GroupGradeLevelShortName = '" & sGroupGradeLevelShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											lGroupGradeLevelID = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
										Else
											lErrorNumber = L_ERR_NO_RECORDS
											sErrorUpload = sErrorUpload & "La clave del GGN no existe.<BR />"
										End If
									End If
								Else
									lGroupGradeLevelID = -1
								End If
							Case "IntegrationID"
								If Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) > 0 Then
									lIntegrationID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								Else
									lIntegrationID = -1
								End If
							Case "JobID"
								lJobID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "JobTypeID"
								sJobTypeShortName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobTypeID From JobTypes Where JobTypeShortName = '" & sJobTypeShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lJobTypeID = CLng(oRecordset.Fields("JobTypeID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del tipo de ocupación no existe.<BR />"
									End If
								End If
							Case "JourneyID"
								lJourneyID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sJourneyShortName = Right(("00" & lJourneyID), Len("00"))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyID From Journeys Where JourneyShortName = '" & sJourneyShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lJourneyID = CLng(oRecordset.Fields("JourneyID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del turno no existe.<BR />"
									End If
								End If
							Case "LevelID"
								If Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) > 0 Then
									lLevelID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									sLevelShortName = Right(("000" & lLevelID), Len("000"))
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select LevelID From Levels Where LevelShortName = '" & sLevelShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											lLevelID = CLng(oRecordset.Fields("LevelID").Value)
										Else
											lErrorNumber = L_ERR_NO_RECORDS
											sErrorUpload = sErrorUpload & "La clave del nivel no existe.<BR />"
										End If
									End If
								Else
									lLevelID = -1
								End If
							Case "PaymentCenterID"
								lPaymentCenterID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sPaymentCenterShortName = Right(("00000" & lPaymentCenterID), Len("00000"))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentCenterID From PaymentCenters Where PaymentCenterShortName = '" & sPaymentCenterShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del centro de pago no existe.<BR />"
									End If
								End If
							Case "PositionID"
								sPositionShotName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "ServiceID"
								lServiceID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If Not IsNumeric(lServiceID) Then
									sServiceShortName = Right(("     " & lServiceID), Len("00000"))
								Else
									sServiceShortName = Right(("00000" & lServiceID), Len("00000"))
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceID From Services Where ServiceShortName = '" & sServiceShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lServiceID = CLng(oRecordset.Fields("ServiceID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del servicio no existe.<BR />"
									End If
								End If
							Case "ShiftID"
								lShiftID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sShiftShortName = Right(("0000" & lShiftID), Len("0000"))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ShiftID From Shifts Where ShiftShortName = '" & sShiftShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lShiftID = CLng(oRecordset.Fields("ShiftID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del horario no existe.<BR />"
									End If
								End If
							Case "StatusID"
								lStatusID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sStatusShortName = CStr(lStatusID)
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID From StatusJobs Where StatusShortName = '" & sStatusShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lStatusID = CLng(oRecordset.Fields("StatusID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del estatus de la plaza no existe.<BR />"
									End If
								End If
							 Case "StartDate"
								Select Case sStartDateFormat
									Case "YYYYMMDD"
										lStartDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lStartDate = CLng(asInputDate(2) & asInputDate(1) & asInputDate(0))
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lStartDate = CLng(asInputDate(2) & asInputDate(0) & asInputDate(1))
								End Select
							Case "EndDate"
								Select Case sEndDateFormat
									Case "YYYYMMDD"
										lEndDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
						End Select
					Next
					If (lErrorNumber = 0) And (lReasonID <> 60) Then
						If (Len(sPositionShotName) > 0) And ((lLevelID > 0) Or ((lGroupGradeLevelID > 0) And (lIntegrationID > 0) And (lClassificationID > 0))) Then
							If lLevelID > 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID From Positions Where (PositionShortName='" & sPositionShotName & "') And (LevelID=" & lLevelID & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID From Positions Where (PositionShortName='" & sPositionShotName & "') And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lPositionID = CLng(oRecordset.Fields("PositionID").Value)
								Else
									lErrorNumber = L_ERR_NO_RECORDS
									sErrorUpload = sErrorUpload & "La clave del puesto no existe.<BR />"
								End If
							End If
						Else
							lErrorNumber = L_ERR_NO_RECORDS
							sErrorUpload = sErrorUpload & "Para obtener la clave del puesto se requiere clave de puesto y clave del nivel o en caso de funcionarios la clave del GGN, la clave de integración y la clave de clasificación.<BR />"
						End If
					End If
					If lErrorNumber = 0 Then
						Select Case lReasonID
							Case 54
								If lErrorNumber = 0 Then
									aJobComponent(N_ID_JOB) = lJobID
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									If Not IsNumeric(lServiceID) Then
										sServiceShortName = Right(("     " & lServiceID), Len("00000"))
									Else
										sServiceShortName = Right(("00000" & lServiceID), Len("00000"))
									End If
									aJobComponent(N_SERVICE_ID_JOB) = CLng(lServiceID)
									If Len(sErrorUpload)= 0 Then
										sErrorUpload = "No se pudo modificar el servicio de la plaza."
									End If
									aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = False
									aJobComponent(B_IS_DUPLICATED_JOB) = False
									aJobComponent(N_JOB_DATE_JOB) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
									lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
								End If
							Case 60
								If lErrorNumber = 0 Then
									Call InitializeJobComponent(oRequest, aJobComponent)
									aJobComponent(N_ID_JOB) = lJobID
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									If lErrorNumber = 0 Then
										If lAreaID <> 0 Then
											aJobComponent(N_AREA_ID_JOB) = lAreaID
											sErrorDescription = "No se pudo obtener la clave de la zona solicitada."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID From Areas Where (AreaID = " & aJobComponent(N_AREA_ID_JOB) & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
											If lErrorNumber = 0 Then
												If Not oRecordset.EOF Then
													aJobComponent(N_ZONE_ID_JOB) = CLng(oRecordset.Fields("ZoneID").Value)
												Else
													aJobComponent(N_ZONE_ID_JOB) = -1
												End If
											Else
												sErrorUpload = sErrorUpload & "No se pudo obtener la zona del área de la plaza " & aJobComponent(S_NUMBER_JOB) & "<BR />"
											End If
										End If
										If lJourneyID <> 0 Then
											aJobComponent(N_JOURNEY_ID_JOB) = lJourneyID
										End If
										If lPaymentCenterID <> 0 Then
											aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
										End If
										If lServiceID <> 0 Then
											aJobComponent(N_SERVICE_ID_JOB) = lServiceID
										End If
										If lShiftID <> 0 Then
											aJobComponent(N_SHIFT_ID_JOB) = lShiftID
										End If
										If lStatusID <> -1 Then
											aJobComponent(N_STATUS_ID_JOB) = lStatusID
										End If
										If lStartDate <> 0 Then
											aJobComponent(N_JOB_DATE_JOB) = lStartDate
										End If
										If lErrorNumber = 0 Then
											aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = False
											aJobComponent(B_IS_DUPLICATED_JOB) = False
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												sErrorUpload = sErrorUpload & "No se pudo modificar la plaza " & aJobComponent(S_NUMBER_JOB) & "<BR />"
											End If
										End If
									Else
										sErrorUpload = sErrorUpload & "No se pudo modificar la plaza " & aJobComponent(N_ID_JOB) & "<BR />"
									End If
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
								End If
							Case 61
								If lErrorNumber = 0 Then
									Call InitializeJobComponent(oRequest, aJobComponent)
									aJobComponent(N_ID_JOB) = CLng(lJobID)
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									If lErrorNumber = 0 Then
										If lPositionID <> 0 Then
											aJobComponent(N_POSITION_ID_JOB) = lPositionID
										End If
										If lLevelID <> 0 Then
											aJobComponent(N_LEVEL_ID_JOB) = lLevelID
										End If
										If lGroupGradeLevelID <> 0 Then
											aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = lGroupGradeLevelID
										End If
										If lIntegrationID <> 0 Then
											aJobComponent(N_INTEGRATION_ID_JOB) = lIntegrationID
										End If
										If lClassificationID <> 0 Then
											aJobComponent(N_CLASSIFICATION_ID_JOB) = lClassificationID
										End If
										If lStartDate <> 0 Then
											aJobComponent(N_JOB_DATE_JOB) = lStartDate
										End If
										If lErrorNumber = 0 Then
											aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = False
											aJobComponent(B_IS_DUPLICATED_JOB) = False
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												sErrorUpload = sErrorUpload & "No se pudo modificar la plaza " & aJobComponent(S_NUMBER_JOB) & "<BR />"
											End If
											If lErrorNumber = 0 Then
												lErrorNumber = UpdateHistoryForward(oADODBConnection, aJobComponent, sErrorDescription)
												If lErrorNumber <> 0 Then 
													sErrorUpload = sErrorUpload & "No se pudo modificar el historial de la plaza " & aJobComponent(S_NUMBER_JOB) & "<BR />"
												End If
											End If
										End If
									Else
										sErrorUpload = sErrorUpload & "No se pudo modificar la plaza " & aJobComponent(N_ID_JOB) & "<BR />"
									End If
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
								End If
							Case Else
								If lErrorNumber = 0 Then
									Call InitializeJobComponent(oRequest, aJobComponent)
									If lReasonID = 59 Then
										If lJobID = 0 Then 
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=100)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=100)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									End If
									If lErrorNumber = 0 Then
										If lReasonID = 59 Then
											If lJobID = 0 Then 
												aJobComponent(N_ID_JOB) = CLng(oRecordset.Fields("CurrentID").Value) + 1
												bAddConsecutive = True
											Else
												bAddConsecutive = False
												aJobComponent(N_ID_JOB) = lJobID
											End If
										Else
											aJobComponent(N_ID_JOB) = CLng(oRecordset.Fields("CurrentID").Value) + 1
										End If
										aJobComponent(S_NUMBER_JOB) = Right(("000000" & aJobComponent(N_ID_JOB)), Len("000000"))
										If lAreaID <> 0 Then
											aJobComponent(N_AREA_ID_JOB) = lAreaID
										End If
										If lJobTypeID <> 0 Then
											aJobComponent(N_JOB_TYPE_ID_JOB) = lJobTypeID
										End If
										If lJourneyID <> 0 Then
											aJobComponent(N_JOURNEY_ID_JOB) = lJourneyID
										End If
										If lPaymentCenterID <> 0 Then
											aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
										End If
										If lPositionID <> 0 Then
											aJobComponent(N_POSITION_ID_JOB) = lPositionID
										End If
										If lServiceID <> 0 Then
											aJobComponent(N_SERVICE_ID_JOB) = lServiceID
										End If
										If lShiftID <> 0 Then
											aJobComponent(N_SHIFT_ID_JOB) = lShiftID
										End If
										If lStartDate <> 0 Then
											aJobComponent(N_START_DATE_JOB) = lStartDate
										End If
										If lEndDate <> 0 Then
											aJobComponent(N_END_DATE_JOB) = lEndDate
										End If
										sErrorDescription = "No se pudo obtener la clave de la zona solicitada."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID From Areas Where (AreaID = " & aJobComponent(N_AREA_ID_JOB) & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												aJobComponent(N_ZONE_ID_JOB) = CLng(oRecordset.Fields("ZoneID").Value)
											End If
										End If
										If lErrorNumber = 0 Then
											aJobComponent(N_ACTIVE_JOB) = 0
											aJobComponent(N_STATUS_ID_JOB) = 2
											lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, bAddConsecutive, sErrorDescription)
											If lErrorNumber <> 0 Then
												sErrorUpload = sErrorUpload & "No se pudo agregar la plaza " & aJobComponent(S_NUMBER_JOB) & "<BR />"
											End If
										End If
									End If
								End If
								If lErrorNumber <> 0 Then
									sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
								End If
						End Select
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If
	UploadJobsFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesConceptFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesConceptsLKP table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesConceptFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sStartDateFormat
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sExecuteQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iDocumentDateYear
	Dim iDocumentDateMonth
	Dim sConceptShortName

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
                        Case "EmployeeID"
                                aiFieldsOrder(0) = iIndex & ",EmployeeID"
                        Case "ConceptID"
                                aiFieldsOrder(1) = iIndex & ",ConceptID"
                        Case "OcurredStartDateYYYYMMDD"
                                sStartDateFormat = "YYYYMMDD"
                                aiFieldsOrder(2) = iIndex & ",StartDate"
                        Case "OcurredStartDateDDMMYYYY"
                                sStartDateFormat = "DDMMYYYY"
                                aiFieldsOrder(2) = iIndex & ",StartDate"
                        Case "OcurredStartDateMMDDYYYY"
                                sStartDateFormat= "MMDDYYYY"
                                aiFieldsOrder(2) = iIndex & ",StartDate"
                        Case "ConceptAmount"
                                aiFieldsOrder(3) = iIndex & ",ConceptAmount"
                        Case "ConceptMin"
                                aiFieldsOrder(4) = iIndex & ",ConceptMin"
                        Case "ConceptMax"
                                aiFieldsOrder(5) = iIndex & ",ConceptMax"
                        Case "AbsenceTypeID"
                                aiFieldsOrder(6) = iIndex & ",AbsenceTypeID"
                        Case "Active"
                                aiFieldsOrder(7) = iIndex & ",Active"
				End Select
			End If
		Next
		sFields = "EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, StartUserID, EndUserID"
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
        If InStr(1, sFields, "EmployeeID") > 0 Then  
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene el Número de empleado."
        ElseIf InStr(1, sFields, "ConceptID") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene el Número de Concepto."
        ElseIf InStr(1, sFields, "StartDate") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene la Fecha de inicio."
        ElseIf InStr(1, sFields, "ConceptAmount") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene el Monto."
        ElseIf InStr(1, sFields, "ConceptMin,", 1) > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene el Mínimo."
        ElseIf InStr(1, sFields, "ConceptMax,") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene el Máximo."
        ElseIf InStr(1, sFields, "AbsenceTypeID") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene Ausencias."
        ElseIf InStr(1, sFields, "Active") > 0 Then
                lErrorNumber = -1
                sErrorDescription = "La información a registrar no contiene si es Activo."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sValues = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sFields, "EndDate", 30000000), "CurrencyID",0), "ConceptQttyID", 1), "ConceptTypeID", 3), "ConceptMinQttyID", 1), "ConceptMaxQttyID", 1), "AppliesToID", -1), "ConceptOrder", 1), "RegistrationDate", sDate), "StartUserID", aLoginComponent(N_USER_ID_LOGIN)), "EndUserID", -1)
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					sQuery = "Insert Into EmployeesConceptsLKP ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					sQuery = sQuery & sFields & ") Values ("
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
							Case "ConceptID"
								sConceptShortName = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sErrorDescription = "No se pudo obtener el identificador del Concepto."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID From Concepts Where ConceptShortName = '" & sConceptShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										sQuery = sQuery & oRecordset.Fields("ConceptID").Value & ", "
									End If
								End If
							Case "StartDate"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
							Case "StartDate"
								Select Case sStartDateFormat
									Case "YYYYMMDD"
										sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(1) & asInputDate(0) & ", "
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										sQuery = sQuery & asInputDate(2) & asInputDate(0) & asInputDate(1) & ", "
								End Select
							Case "ConceptAmount"
								sQuery = sQuery & asFileRow(CDbl(aiFieldsOrder(jIndex)(0))) & ", "
							Case "ConceptMin"
								sQuery = sQuery & asFileRow(CDbl(aiFieldsOrder(jIndex)(0))) & ", "
							Case "ConceptMax"
								sQuery = sQuery & asFileRow(CDbl(aiFieldsOrder(jIndex)(0))) & ", "
							Case "AbsenceTypeID"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
							Case "Active"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "						
						End Select
					Next
					sQuery = sQuery & sValues & ")"
					sErrorDescription = "No se pudo guardar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If
	UploadEmployeesConceptFile = lErrorNumber
	Err.Clear
End Function

Function UploadEmployeesFeaturesFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesAbsencesLKP table.
'Inputs:  sAction, oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadEmployeesFeaturesFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sStartDateFormat
	Dim sEndDateFormat
	Dim sPayrollDateFormat
	Dim asInputDate
	Dim sFields
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim bValidConceptKey
	Dim iConceptID
	Dim sQueryConcept
	Dim sEmployeeID
	Dim sStartDate
	Dim lStartDate
	Dim sStartDateNightShifts
	Dim sEndDate
	Dim lEndDate
	Dim sPayrollDate
	Dim lPayrollDate
	Dim iConceptQttyID
	Dim sAppliesToID
	Dim iAmount
	Dim iStartHour
	Dim iEndHour
	Dim sMessage
	Dim iEmployeeTypeID
	Dim iPositionTypeID
	Dim iJourneyID
	Dim iStatusID
	Dim sNightShiftDates
	Dim iNightShifts
	Dim sEmployeeGrade

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item				
					Case "EmployeeID"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeID"
					Case "StartDateYYYYMMDD"
						sStartDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "StartDateDDMMYYYY"
						sStartDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "StartDateMMDDYYYY"
						sStartDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",StartDate"
					Case "EndDateYYYYMMDD"
						sEndDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateDDMMYYYY"
						sEndDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "EndDateMMDDYYYY"
						sEndDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",EndDate"
					Case "PayrollDateYYYYMMDD"
						sPayrollDateFormat = "YYYYMMDD"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateDDMMYYYY"
						sPayrollDateFormat = "DDMMYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "PayrollDateMMDDYYYY"
						sPayrollDateFormat = "MMDDYYYY"
						aiFieldsOrder(iIndex) = iIndex & ",PayrollDate"
					Case "ConceptAmount"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptAmount"
					Case "ConceptComments"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptComments"
					Case "StartHour3"
						aiFieldsOrder(iIndex) = iIndex & ",StartHour3"
					Case "EndHour3"
						aiFieldsOrder(iIndex) = iIndex & ",EndHour3"
					Case "ConceptQttyID"
						aiFieldsOrder(iIndex) = iIndex & ",ConceptQttyID"
					Case "YearID"
						aiFieldsOrder(iIndex) = iIndex & ",YearID"
					Case "EmployeeGrade"
						aiFieldsOrder(iIndex) = iIndex & ",EmployeeGrade"
				End Select
			End If
		Next

		Select Case lReasonID
			Case EMPLOYEES_ANUAL_AWARD, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_EXCENT, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_GLASSES, EMPLOYEES_MONTHAWARD, EMPLOYEES_MOTHERAWARD, EMPLOYEES_NON_EXCENT, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_FONAC_ADJUSTMENT
				sFields = "EmployeeID, PayrollDate, ConceptAmount, "
			Case EMPLOYEES_FONAC_CONCEPT
				 sFields = "EmployeeID, PayrollDate, "
			Case EMPLOYEES_NIGHTSHIFTS
				sFields = "EmployeeID, StartDate, PayrollDate, "
			Case EMPLOYEES_EFFICIENCY_AWARD
				sFields = "EmployeeID, ConceptAmount, "
			Case EMPLOYEES_GRADE
				sFields = "EmployeeID, YearID, PayrollDate, EmployeeGrade, "
			Case Else
				sFields = "EmployeeID, StartDate, EndDate, PayrollDate, ConceptAmount, "
		End Select

		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el número de empleado."
		ElseIf InStr(1, sFields, "StartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de inicio."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de fin."
		ElseIf InStr(1, sFields, "PayrollDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha nómina."
		ElseIf InStr(1, sFields, "YearID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el año para la calificación."
		ElseIf InStr(1, sFields, "EmployeeGrade") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la calificación del empleado."
		ElseIf (InStr(1, sFields, "ConceptAmount") > 0) And (lReasonID <> EMPLOYEES_HELP_COMISSION) And (lReasonID <> EMPLOYEES_CONCEPT_08) And (lReasonID <> EMPLOYEES_ADDITIONALSHIFT) Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el importe de la prestación."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			sErrorQueries = ""
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					lErrorNumber = 0
					sEndDate = Empty
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "EmployeeID"
								iConceptQttyID = 1
								sAppliesToID = ""
								sEmployeeID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								aEmployeeComponent(N_ID_EMPLOYEE) = sEmployeeID
								sQuery = sQuery & sEmployeeID & ", "
								sErrorDescription = "No existe el empleado indicado"
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE), "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									iEmployeeTypeID = oRecordset.Fields("EmployeeTypeID").Value
									iPositionTypeID = oRecordset.Fields("PositionTypeID").Value
									iJourneyID = oRecordset.Fields("JourneyID").Value
									iStatusID = oRecordset.Fields("StatusID").Value
								End If
								Select Case lReasonID
									Case -89
										iConceptID = 100
									Case EMPLOYEES_CONCEPT_08
										iConceptID = 8
										iConceptQttyID = 2
									Case EMPLOYEES_FOR_RISK
										iConceptID = 4
										iConceptQttyID = 2
										sAppliesToID = "1"
									Case EMPLOYEES_ANTIQUITIES
										iConceptID = 5
										iConceptQttyID = 1
										sAppliesToID = "-1"
									Case EMPLOYEES_ADDITIONALSHIFT
										iConceptID = 7
										sAppliesToID = "1,5"
									Case EMPLOYEES_CONCEPT_08
										iConceptID = 8
										sAppliesToID = "1,5"
									Case EMPLOYEES_CONCEPT_16
										iConceptID = 19
									Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
										iConceptID = 22
									Case EMPLOYEES_GLASSES
										iConceptID = 24
									Case EMPLOYEES_MOTHERAWARD
										iConceptID = 26
									Case EMPLOYEES_ANUAL_AWARD
										iConceptID = 32
									Case EMPLOYEES_FAMILY_DEATH
										iConceptID = 45
									Case EMPLOYEES_PROFESSIONAL_DEGREE
										iConceptID = 46
									Case EMPLOYEES_MONTHAWARD
										iConceptID = 50
										iConceptQttyID = 2
										sAppliesToID = "1"
										iAmount = 20
									Case EMPLOYEES_HELP_COMISSION
										iConceptID = 63
									Case EMPLOYEES_FONAC_CONCEPT
										iConceptID = 77
									Case EMPLOYEES_FONAC_ADJUSTMENT
										iConceptID = 76
									Case EMPLOYEES_SPORTS_HELP
										iConceptID = 165
									Case EMPLOYEES_SPORTS
										iConceptID = 69
									Case EMPLOYEES_BENEFICIARIES
										iConceptID = 70
									Case EMPLOYEES_NON_EXCENT
										iConceptID = 72
									Case EMPLOYEES_EXCENT
										iConceptID = 73
									Case EMPLOYEES_CARLOAN
										iConceptID = 74
									Case EMPLOYEES_BENEFICIARIES_DEBIT
										iConceptID = 86
									Case EMPLOYEES_ADD_SAFE_SEPARATION
										iConceptID = 87
									Case EMPLOYEES_SAFE_SEPARATION
										iConceptID = 120
										iConceptQttyID = 2
										sAppliesToID = "1,3"
									Case EMPLOYEES_NIGHTSHIFTS
										iConceptID = 93
										iAmount = 1/6
									Case EMPLOYEES_CONCEPT_C3
										iConceptID = 94
									Case EMPLOYEES_LICENSES
										iConceptID = 104
									Case EMPLOYEES_EFFICIENCY_AWARD
										iConceptID = 32
										lStartDate = CLng(oRequest("PayrollDateYYYYMMDD").Item)
										lPayrollDate = CLng(oRequest("PayrollDateYYYYMMDD").Item)
								End Select
							Case "StartDate"
								If lErrorNumber = 0 Then
									Select Case lReasonID
										Case EMPLOYEES_NIGHTSHIFTS
											iNightShifts = 0
											Select Case sStartDateFormat
												Case "YYYYMMDD"
													asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
													sStartDateNightShifts = Split(asInputDate, ",")
													For kIndex = 0 To UBound(sStartDateNightShifts)
														iNightShifts = iNightShifts + 1
														sStartDate = sStartDateNightShifts(kIndex)
														If VerifyIfUploadMonthDateIsCorrect(sStartDate, sErrorDescription) Then
															If Not IsHoliday(oADODBConnection, sStartDate, sErrorDescription) Then
																sErrorDescription = "El concepto solo se puede registrar en día festivo y la fecha " & DisplayNumericDateFromSerialNumber(sStartDate) & " no lo es."
																lErrorNumber = -1
															Else
																sNightShiftDates = sNightShiftDates & CStr(sStartDate) & ","
															End If
														Else
															lErrorNumber = -1
														End If
														If (lErrorNumber <> 0) Then Exit For
													Next
												Case "DDMMYYYY"
													asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
													sStartDateNightShifts = Split(asInputDate, ",")
													For kIndex = 0 To UBound(sStartDateNightShifts)
														iNightShifts = iNightShifts + 1
														asInputDate = Split(sStartDateNightShifts(kIndex), Mid(sStartDateNightShifts(kIndex), Len("000"), Len("-")))
														sStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
														lStartDate = CLng(sStartDate)
														If (Err.Number <> 0) Then
															Err.Clear
															sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
															lErrorNumber = -1
														Else
															If VerifyIfUploadMonthDateIsCorrect(sStartDate, sErrorDescription) Then
																If Not IsHoliday(oADODBConnection, sStartDate, sErrorDescription) Then
																	sErrorDescription = "El concepto solo se puede registrar en día festivo y la fecha " & DisplayNumericDateFromSerialNumber(sStartDate) & " no lo es."
																	lErrorNumber = -1
																Else
																	sNightShiftDates = sNightShiftDates & CStr(sStartDate) & ","
																End If
															Else
																lErrorNumber = -1
															End If
														End If
													Next
												Case "MMDDYYYY"
													asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
													sStartDateNightShifts = Split(asInputDate, ",")
													For kIndex = 0 To UBound(sStartDateNightShifts)
														iNightShifts = iNightShifts + 1
														asInputDate = Split(sStartDateNightShifts(kIndex), Mid(sStartDateNightShifts(kIndex), Len("000"), Len("-")))
														sStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
														lStartDate = CLng(sStartDate)
														If (Err.Number <> 0) Then
															Err.Clear
															sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
															lErrorNumber = -1
														Else
															If VerifyIfUploadMonthDateIsCorrect(sStartDate, sErrorDescription) Then
																If Not IsHoliday(oADODBConnection, sStartDate, sErrorDescription) Then
																	sErrorDescription = "El concepto solo se puede registrar en día festivo y la fecha " & DisplayNumericDateFromSerialNumber(sStartDate) & " no lo es."
																	lErrorNumber = -1
																Else
																	sNightShiftDates = sNightShiftDates & CStr(sStartDate) & ","
																End If
															Else
																lErrorNumber = -1
															End If
														End If
													Next
											End Select
											If lErrorNumber = 0 Then
												If InStr(1, Right(sNightShiftDates, Len(",")), ",") Then
													sNightShiftDates = Left(sNightShiftDates, (Len(sNightShiftDates) - Len(",")))
												End If
												iAmount = iNightShifts
											End If
										Case Else
											Select Case sStartDateFormat
												Case "YYYYMMDD"
													sStartDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
												Case "DDMMYYYY"
													asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
													asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
													sStartDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
												Case "MMDDYYYY"
													asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
													asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
													sStartDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
											End Select
											lStartDate = CLng(sStartDate)
											If (Err.Number <> 0) Then
												Err.Clear
												sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
												lErrorNumber = -1
											Else
												If Not VerifyIfUploadMonthDateIsCorrect(lStartDate, sErrorDescription) Then
													lErrorNumber = -1
												End If
											End If
									End Select
								End If
							Case "EndDate"
								If lErrorNumber = 0 Then
									Select Case sEndDateFormat
										Case "YYYYMMDD"
											sEndDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									If Not IsEmpty(sEndDate) Then
										lEndDate = CLng(sEndDate)
									Else
										lEndDate = 0
										Err.Clear
									End If
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la fecha de fin en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(lEndDate, sErrorDescription) Then
											lErrorNumber = -1
										Else
											If (lEndDate = 0) And ((lReasonID <> EMPLOYEES_ANUAL_AWARD) And (lReasonID <> EMPLOYEES_CHILDREN_SCHOOLARSHIPS) And (lReasonID <> EMPLOYEES_CONCEPT_C3) And (lReasonID <> EMPLOYEES_EXCENT) And (lReasonID <> EMPLOYEES_FAMILY_DEATH) And (lReasonID <> EMPLOYEES_GLASSES) And (lReasonID <> EMPLOYEES_MONTHAWARD) And (lReasonID <> EMPLOYEES_MOTHERAWARD) And (lReasonID <> EMPLOYEES_NIGHTSHIFTS) And (lReasonID <> EMPLOYEES_NON_EXCENT) And (lReasonID <> EMPLOYEES_PROFESSIONAL_DEGREE)) Then lEndDate = 30000000
											If lEndDate < lStartDate Then
												sErrorDescription = "La fecha de fin debe de ser mayor a la fecha de inicio"
												lErrorNumber = -1
											End If
										End If
									End If
								End If
							Case "PayrollDate"
								If lErrorNumber = 0 Then
									Select Case sPayrollDateFormat
										Case "YYYYMMDD"
											sPayrollDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										Case "DDMMYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
										Case "MMDDYYYY"
											asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
											sPayrollDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
									End Select
									lPayrollDate = CLng(sPayrollDate)
									If (Err.Number <> 0) Then
										Err.Clear
										sErrorDescription = "Introduzca la quincena de aplicación en un formato correcto."
										lErrorNumber = -1
									Else
										If Not VerifyIfUploadMonthDateIsCorrect(lPayrollDate, sErrorDescription) Then
											lErrorNumber = -1
										Else
											If Not VerifyPayrollIsActive(oADODBConnection, lPayrollDate, N_PAYROLL_FOR_FEATURES, sErrorDescription) Then
												lErrorNumber = -1
											End If
										End If
									End If
								End If
							Case "ConceptAmount"
								If lErrorNumber = 0 Then
									iAmount =  CDbl(asFileRow(CLng(aiFieldsOrder(jIndex)(0))))
									Select Case lReasonID ' Validacion de cantidad por concepto
										Case EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT
											aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 46.1538
										Case EMPLOYEES_GLASSES
											If (iAmount > 1400) Then
												sErrorDescription = "el valor del Importe para ayuda de anteojos no puede ser mayor a $ 1400"
												lErrorNumber = -1
											Else
												sQuery = sQuery & asFileRow(CLng(aiFieldsOrder(jIndex)(0))) & ", "
											End If
										Case EMPLOYEES_FOR_RISK
											If (iAmount <> 10) And (iAmount <> 20) Then
												sErrorDescription = "el porcentaje de la compensación solo puede ser 10 o 20 %"
												lErrorNumber = -1
											Else
												sQuery = sQuery & asFileRow(CLng(aiFieldsOrder(jIndex)(0))) & ", "
											End If
										Case EMPLOYEES_SPORTS
											If (iAmount <> 0) Then
												sErrorDescription = "el valor del Importe para cuota deportivo debe ser 0"
												lErrorNumber = -1
											Else
												sQuery = sQuery & asFileRow(CLng(aiFieldsOrder(jIndex)(0))) & ", "
											End If
										Case Else
											sQuery = sQuery & asFileRow(CLng(aiFieldsOrder(jIndex)(0))) & ", "
									End Select
								End If
							Case "ConceptComments"
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = asFileRow(CLng(aiFieldsOrder(jIndex)(0)))
							Case "StartHour3"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
								iStartHour = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EndHour3"
								sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
								iEndHour = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ConceptQttyID"
								iConceptQttyID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "EmployeeGrade"
								If lErrorNumber = 0 Then
									sEmployeeGrade = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Select Case sEmployeeGrade
										Case "A"
											aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) = 40
										Case "B"
											aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) = 30
										Case "C"
											aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) = 25
										Case "D"
											aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) = 20
										Case "E"
											aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) = 15
										Case Else
											lErrorNumber = -1
											sErrorDescription = "el valor de la calificación para el empleado solamente puede ser A, B, C, D o E"
									End Select
									If lErrorNumber = 0 Then
										aEmployeeComponent(S_EMPLOYEE_GRADE) = sEmployeeGrade
									End If
								End If
							Case "YearID"
								aEmployeeComponent(N_CALIFICATION_YEAR) = CInt(asFileRow(CLng(aiFieldsOrder(jIndex)(0))))
						End Select
					Next
					Select Case lReasonID
						Case EMPLOYEES_ANUAL_AWARD, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_EXCENT, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_GLASSES, EMPLOYEES_MONTHAWARD, EMPLOYEES_MOTHERAWARD, EMPLOYEES_NON_EXCENT, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT
							lStartDate = lPayrollDate
					End Select
					Select Case lReasonID
						Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
							iAmount = 1
					End Select
					If lErrorNumber = 0 Then
						If lReasonID = EMPLOYEES_GRADE Then
							lErrorNumber = AddEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								sErrorDescription = "No se pudo guardar la información de la calificación del empleado, debido a que " & sErrorDescription
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						Else
							If Not VerifyRequerimentsForEmployeesConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = -1
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo obtener la información del empleado."
								'lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = lStartDate
									aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = lEndDate
									aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = lPayrollDate
									aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(iConceptID)
									aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CLng(iConceptQttyID)
									aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = sAppliesToID
									If lReasonID = EMPLOYEES_NIGHTSHIFTS Then
										aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(iAmount / 6)
										aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = sNightShiftDates
									Else
										aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = iAmount
									End If
									aEmployeeComponent(N_START_DATE2_EMPLOYEE) = CLng(sDate)
									aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sFileName
									aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 1
									Select Case lReasonID
										Case -89 sMessage = "devoluciones no gravables."
										Case 53  sMessage = "compensación por riesgos profesionales."
										Case EMPLOYEES_ANTIQUITIES sMessage = "compensación por antigüedades."
										Case EMPLOYEES_ADDITIONALSHIFT sMessage = "turno opcional."
										Case EMPLOYEES_CONCEPT_08 sMessage = "percepcion adicional."
										Case EMPLOYEES_CONCEPT_16 sMessage = "devolución por deducciones indebidas."
										Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS sMessage = "importe de la beca de hijos de trabajadores."
										Case EMPLOYEES_GLASSES sMessage = "ayuda de anteojos."
										Case EMPLOYEES_ANUAL_AWARD sMessage = "importe del estimulo."
										Case EMPLOYEES_FAMILY_DEATH sMessage = "ayuda por muerte de familiar en primer grado."
										Case EMPLOYEES_PROFESSIONAL_DEGREE sMessage = "ayuda por impresión de tesis."
										Case EMPLOYEES_MONTHAWARD sMessage = "premio al trabajador del mes."
										Case EMPLOYEES_SPORTS_HELP sMessage = "apoyo al deporte."
										Case EMPLOYEES_SPORTS sMessage = "cuota deportiva."
										Case EMPLOYEES_BENEFICIARIES sMessage = "importe para pensión alimenticia."
										Case EMPLOYEES_NON_EXCENT sMessage = "deducción por cobro de sueldos indebidos."
										Case EMPLOYEES_EXCENT sMessage = "importe para otras deducciones."
										Case EMPLOYEES_MOTHERAWARD sMessage = "importe del Premio del 10 de Mayo."
										Case EMPLOYEES_BENEFICIARIES_DEBIT sMessage = "importe para adeudo pensión alimenticia."
										Case EMPLOYEES_ADD_SAFE_SEPARATION sMessage = "importe para seguro adicional de separación individualizado para el personal de mando."
										Case EMPLOYEES_NIGHTSHIFTS sMessage = "jornada nocturna adicional."
										Case EMPLOYEES_CONCEPT_C3 sMessage = "premios, estímulos y recompensas."
										Case EMPLOYEES_LICENSES sMessage = "importe de retenciones por exceso de licencias médicas."
										Case EMPLOYEES_SAFE_SEPARATION sMessage = "seguro de separación individualizado para personal de mando."
										Case EMPLOYEES_CARLOAN sMessage = "préstamo automóvil servidores públicos de mando superior."
										Case EMPLOYEES_EFFICIENCY_AWARD = "importe del estímulo a la productividad, calidad y eficacia."
									End Select
									aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
									If lReasonID = EMPLOYEES_SAFE_SEPARATION Then
										If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 4 And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 5) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 10) Then
											lErrorNumber = -1
											sErrorDescription = "El porcentaje para SI solo puede ser 2, 4, 5 y 10."
										End If
										If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 1 Then
											lErrorNumber = -1
											sErrorDescription = "El SI solo puede otorgarse a funcionarios"
										End If
									ElseIf lReasonID = EMPLOYEES_ADD_SAFE_SEPARATION Then
										sErrorDescription = "No se pudo obtener la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=120) And (StartDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate<=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ") And (Active=1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If oRecordset.EOF Then
												sErrorDescription = "Para registrar el seguro adicional, debe de estar registrado el concepto SI."
												lErrorNumber = -1
											End If
											oRecordset.Close
										End If
										If (aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 100) Then
											lErrorNumber = -1
											sErrorDescription = "El porcentaje para AE no puede ser mayor al 100 %."
										End If
										If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 1 Then
											lErrorNumber = -1
											sErrorDescription = "El AE solo puede otorgarse a funcionarios"
										End If
									ElseIf lReasonID = EMPLOYEES_ADDITIONALSHIFT Then
										If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 1 Then
											sErrorDescription = "el empleado no tiene puesto de base."
											lErrorNumber = -1
										End If
									ElseIf lReasonID = EMPLOYEES_CONCEPT_08 Then
										If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1 Then
											sErrorDescription = "el empleado tiene tabulador de funcionario."
											lErrorNumber = -1
										Else
											If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 2 Then
												sErrorDescription = "el empleado no tiene puesto de confianza."
												lErrorNumber = -1
											End If
										End If
									End If
									If lErrorNumber = 0 Then
										aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
										lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
									If lErrorNumber <> 0 Then
										sErrorDescription = "No se pudo guardar la información del registro de " & sMessage & " Debido a que " & sErrorDescription
										sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
									End If
								End If
							Else
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						End If
					Else
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					End If
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If
	UploadEmployeesFeaturesFile = lErrorNumber
	Err.Clear
End Function

Function UploadFONACFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         EmployeesFONAC table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadFONACFile"
	Dim sFileContents
	Dim asFileContents
	Dim iIndex
	Dim lAreaID
	Dim lPositionTypeID
	Dim lPayrollID
	Dim sErrorQueries
	Dim lErrorNumber

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		sErrorQueries = ""
		For iIndex = 0 To UBound(asFileContents)
			If Len(asFileContents(iIndex)) = 110 Then
				lAreaID = -1
				Call GetNameFromTable(oADODBConnection, "AreasFromCodes", "'" & Mid(asFileContents(iIndex), 68, 5) & "'", "", "", lAreaID, "")
				If StrComp(Mid(asFileContents(iIndex), 104, 1), "C", vbBinaryCompare) = 0 Then
					lPositionTypeID = 2
				Else
					lPositionTypeID = 1
				End If
				lPayrollID = GetPayrollFromNumber(Mid(asFileContents(iIndex), 84, 4) & Mid(asFileContents(iIndex), 82, 2))
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesFONAC (EmployeeID, EmployeeNumber, EmployeeName, EmployeeRFC, AreaID, PositionTypeID, PayrollID, Concept77Amount, Concept54Amount) Values (" & Left(asFileContents(iIndex), 6) & ", '" & Trim(Left(asFileContents(iIndex), 6)) & "', '" & Trim(Mid(asFileContents(iIndex), 17, 50)) & "', '" & Trim(Mid(asFileContents(iIndex), 7, 10)) & "', " & lAreaID & ", " & lPositionTypeID & ", " & lPayrollID & ", " & Mid(asFileContents(iIndex), 98, 4) & "." & Mid(asFileContents(iIndex), 102, 2) & ", " & Mid(asFileContents(iIndex), 105, 4) & "." & Mid(asFileContents(iIndex), 109, 2) & ")", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber <> 0 Then
					sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					lErrorNumber = 0
				End If
			ElseIf Len(asFileContents(iIndex)) > 0 Then
				sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>No contiene los caracteres necesarios.<BR /><BR />"
			End If
		Next
		sErrorDescription = sErrorQueries
	End If

	UploadFONACFile = lErrorNumber
	Err.Clear
End Function

Function UploadMedicalAreasFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         MedicalAreas table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadMedicalAreasFile"
	
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormatDocument
	Dim sDateFormatStart
	Dim sDateFormatEnd
	Dim asInputDate
	Dim sFields
	Dim sValues
	Dim sQuery
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iDocumentDateYear
	Dim iDocumentDateMonth
	Dim iServiceIDLen
	
	
	sErrorDescription = "No se pudo eliminar la matriz UNIMED cargada en el sistema"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From MedicalAreas", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		sFileContents = GetFileContents(sFileName, sErrorDescription)
		If Len(sFileContents) > 0 Then
			asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
			asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
			aiFieldsOrder = ""
			aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
			For iIndex = 0 To UBound(asFileRow)
				If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
				ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
				Else
					Select Case oRequest("Column" & (iIndex + 1)).Item
						Case "MedicalAreasID"
							aiFieldsOrder(0) = iIndex & ",MedicalAreasID"
						Case "CompanyID"
							aiFieldsOrder(1) = iIndex & ",CompanyID"
						Case "MedicalAreasTypeID"
							aiFieldsOrder(2) = iIndex & ",MedicalAreasTypeID"
						Case "PositionID"
							aiFieldsOrder(3) = iIndex & ",PositionID"
						Case "ServiceID"
							aiFieldsOrder(4) = iIndex & ",ServiceID"
						Case "ColumnNumber"
							aiFieldsOrder(5) = iIndex & ",ColumnNumber"
					End Select
				End If
			Next

			sFields = "MedicalAreasID, CompanyID, MedicalAreasTypeID, PositionID, ServiceID, ColumnNumber, "
			For iIndex = 0 To UBound(aiFieldsOrder)
				aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
				If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
			Next

			If InStr(1, sFields, "MedicalAreasID") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene el número de fila."
			ElseIf InStr(1, sFields, "CompanyID") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene la clave de la empresa."
			ElseIf InStr(1, sFields, "MedicalAreasTypeID") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene el tipo de reporte UNIMED."
			ElseIf InStr(1, sFields, "PositionID") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene la clave de puesto."
			ElseIf InStr(1, sFields, "ServiceID") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene la clave del servicio."
			ElseIf InStr(1, sFields, "ColumnNumber") > 0 Then
				lErrorNumber = -1
				sErrorDescription = "La información a registrar no contiene el número del anexo."
			Else
				sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
				sValues = Replace(sFields,"ColumnNumber,","ColumnNumber")
				sErrorQueries = ""
				For iIndex = 0 To UBound(asFileContents)
					If Len(asFileContents(iIndex)) > 0 Then
						asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)

						sQuery = "Insert Into MedicalAreas ("
						For jIndex = 0 To UBound(aiFieldsOrder)
							If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
						Next
						sQuery = Left(sQuery, Len(sQuery)-2)
						sQuery = sQuery & sFields & ") Values ("
						sErrorDescription = "No se pudo guardar la información del registro."
						For jIndex = 0 To UBound(aiFieldsOrder)
							Select Case aiFieldsOrder(jIndex)(1)
								Case "PositionID"
									If Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))> 4 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID From Positions Where (PositionShortName = '" & Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´") & "')", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												sQuery = sQuery & CStr(oRecordset.Fields("PositionID").Value) & ", "
											Else
												sErrorDescription = "No existe la clave del puesto."
												lErrorNumber = iIndex + 1
											End If
										End If
									Else
										sQuery = sQuery & "-1, "
									End If
								Case "ServiceID"
									iServiceIDLen = Len(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									If iServiceIDLen > 3 Then
										Select Case iServiceIDLen
											Case 3 
												asFileRow(CInt(aiFieldsOrder(jIndex)(0))) = "00" & asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
											Case 4
												asFileRow(CInt(aiFieldsOrder(jIndex)(0))) = "0" & asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										End Select
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceID From Services Where (ServiceShortName = '" & Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´") & "')", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												sQuery = sQuery & CStr(oRecordset.Fields("ServiceID").Value) & ", "
											Else
												sErrorDescription = "No existe la clave del servicio."
												lErrorNumber = iIndex + 1
											End If
										End If
									Else
										sQuery = sQuery & "-1, "
									End If
								Case "ColumnNumber"
									sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & " "
								Case ""
								Case Else
									sQuery = sQuery & asFileRow(CInt(aiFieldsOrder(jIndex)(0))) & ", "
							End Select
						Next
						sQuery = sQuery & sValues & ")"
						If lErrorNumber = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						Else
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						End If
					End If
				Next
			End If
			If Len(sErrorQueries) > 0 Then
				lErrorNumber = -1
				sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
			End If
		End If
	End If

	UploadMedicalAreasFile = lErrorNumber
	Err.Clear
End Function

Function UploadProfessionalRiskFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         Professional Risk Matrix.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadProfessionalRiskFile"
	Dim sFileContents
	Dim asFileContents
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim lErrorNumberUpload
	Dim asFileRow
	Dim oRecordset
	Dim sQuery
	Dim sQueryPosition
	Dim sErrorUpload
	Dim lPositionColumn
	
	sQuery = "Truncate Table ProfessionalRiskMatrix"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		sErrorUpload = ""
		For iIndex = 0 To UBound(asFileContents)
			asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
			If iIndex = 0 Then
				For jIndex = 1 To UBound(asFileRow)
					If StrComp(oRequest("Column" & (jIndex)).Item, "PositionID", vbBinaryCompare) = 0 Then 
						lPositionColumn = jIndex - 1
						jIndex = UBound(asFileRow)
					End If
				Next
			End If

			sQueryPosition = "Select PositionID From Positions Where PositionShortName Like '" & asFileRow(lPositionColumn) & "'"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryPosition, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If oRecordset.EOF Then
				lErrorNumberUpload = L_ERR_NO_RECORDS
				sErrorUpload = sErrorUpload & "El puesto " & asFileRow(lPositionColumn) & " no existe.<BR />"
			Else
				Do While Not oRecordset.EOF
					sQuery = "Insert Into ProfessionalRiskMatrix ("
					For jIndex = 0 To UBound(asFileRow)
						If jIndex < UBound(asFileRow) Then
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ","
						Else
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ",ModifyDate, UserID, Active) Values("
						End If
					Next

					For jIndex = 0 To UBound(asFileRow)
						If jIndex = lPositionColumn Then
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & oRecordset.Fields("PositionID").Value & ","
							Else
								sQuery = sQuery & oRecordset.Fields("PositionID").Value & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1)"
							End If
						Else
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & asFileRow(jIndex) & ","
							Else
								sQuery = sQuery & asFileRow(jIndex) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1)"
							End If
						End If
					Next
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					oRecordset.MoveNext
				Loop
			End If
		Next
		oRecordset.Close
		Set oRecorset = Nothing
		sErrorDescription = sErrorUpload
	End If
End Function

Function UploadPositionsSpecialJourneysLKPFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         ConceptValues table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadPositionsSpecialJourneysLKPFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim sDateFormat
	Dim asInputDate
	Dim sFields
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sErrorQueries
	Dim iTotalRecord
	Dim lStartDateForValueConcept
	iTotalRecord = 0

	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case oRequest("Column" & (iIndex + 1)).Item
					Case "PositionShortName"
						aiFieldsOrder(0) = iIndex & ",PositionShortName"
					Case "LevelID"
						aiFieldsOrder(1) = iIndex & ",LevelID"
					Case "WorkingHours"
						aiFieldsOrder(2) = iIndex & ",WorkingHours"
					Case "ServiceID"
						aiFieldsOrder(3) = iIndex & ",ServiceID"
					Case "CenterTypeID"
						aiFieldsOrder(4) = iIndex & ",CenterTypeID"
					Case "IsActive1"
						aiFieldsOrder(5) = iIndex & ",IsActive1"
					Case "IsActive2"
						aiFieldsOrder(6) = iIndex & ",IsActive2"
					Case "IsActive3"
						aiFieldsOrder(7) = iIndex & ",IsActive3"
					Case "IsActive4"
						aiFieldsOrder(8) = iIndex & ",IsActive4"
					Case "StartDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(9) = iIndex & ",StartDate"
					Case "StartDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(9) = iIndex & ",StartDate"
					Case "StartDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(9) = iIndex & ",StartDate"
					Case "EndDateYYYYMMDD"
						sDateFormat = "YYYYMMDD"
						aiFieldsOrder(10) = iIndex & ",EndDate"
					Case "OcurredEndDateDDMMYYYY"
						sDateFormat = "DDMMYYYY"
						aiFieldsOrder(10) = iIndex & ",EndDate"
					Case "OcurredEndDateMMDDYYYY"
						sDateFormat = "MMDDYYYY"
						aiFieldsOrder(10) = iIndex & ",EndDate"
				End Select
			End If
		Next
		sFields = "PositionShortName, LevelID, WorkingHours, ServiceID, CenterTypeID, StartDate, IsActive1, IsActive2, IsActive3, IsActive4, "
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "PositionShortName") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el puesto."
		ElseIf InStr(1, sFields, "LevelID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el nivel del puesto."
		ElseIf InStr(1, sFields, "WorkingHours") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene las horas de trabajo (Jornada)."
		ElseIf InStr(1, sFields, "ServiceID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el servicio."
		ElseIf InStr(1, sFields, "CenterTypeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el tipo de centro de trabajo."
		ElseIf InStr(1, sFields, "IsActive1") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el valor que indica si aplica para guardias."
		ElseIf InStr(1, sFields, "IsActive2") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el valor que indica si aplica para suplencias."
		ElseIf InStr(1, sFields, "IsActive3") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el valor que indica si aplica para rezago q."
		ElseIf InStr(1, sFields, "IsActive4") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el valor que indica si aplica para PROVAC."
		ElseIf InStr(1, sFields, "StartDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de inicio."
		ElseIf InStr(1, sFields, "EndDate") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene la fecha de termino."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				If Len(asFileContents(iIndex)) > 0 Then
					lErrorNumber = 0
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					aConceptComponent(N_RECORD_ID_CONCEPT) = -1
					aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
					aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1
					aConceptComponent(N_SERVICE_ID_CONCEPT) = -1
					aConceptComponent(N_CENTER_TYPE_ID) = -1
					aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
					aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT) = True
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "PositionShortName"
								aConceptComponent(S_SHORT_NAME_CONCEPT) = CStr(Replace(asFileRow(CInt(aiFieldsOrder(jIndex)(0))), "'", "´"))
							Case "LevelID"
								aConceptComponent(N_LEVEL_ID_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select LevelID From Levels Where (LevelShortName='" & Right(("00" & aConceptComponent(N_LEVEL_ID_CONCEPT)), Len("000")) & "') And (EndDate=30000000) And (Active=1)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aConceptComponent(N_LEVEL_ID_CONCEPT) = CInt(oRecordset.Fields("LevelID").Value)
									Else
										lErrorNumber = -1
										sErrorDescription = "No existe el nivel de la clave especificada " & aConceptComponent(N_LEVEL_ID_CONCEPT)
									End If
								Else
									lErrorNumber = -1
									sErrorDescription = "No se pudo obtener el nivel de la clave especificada " & aConceptComponent(N_LEVEL_ID_CONCEPT)
								End If
							Case "WorkingHours"
								aConceptComponent(D_WORKING_HOURS_CONCEPT) = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
							Case "ServiceID"
								If lErrorNumber = 0 Then
									aConceptComponent(N_SERVICE_ID_CONCEPT) = (asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceID From Services Where (ServiceShortName='" & Right(("0000" & aConceptComponent(N_SERVICE_ID_CONCEPT)), Len("00000")) & "') And (EndDate=30000000) And (Active=1)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											aConceptComponent(N_SERVICE_ID_CONCEPT) = CInt(oRecordset.Fields("ServiceID").Value)
										Else
											lErrorNumber = -1
											sErrorDescription = "No existe el servicio de la clave especificada " & aConceptComponent(N_SERVICE_ID_CONCEPT)
										End If
									Else
										lErrorNumber = -1
										sErrorDescription = "No se pudo obtener el servicio de la clave especificada " & aConceptComponent(N_SERVICE_ID_CONCEPT)
									End If
								End If
							Case "CenterTypeID"								
								If lErrorNumber = 0 Then
									aConceptComponent(N_CENTER_TYPE_ID) = (asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CenterTypeID From CenterTypes Where (CenterTypeShortName='" & Right(("00" & aConceptComponent(N_CENTER_TYPE_ID)), Len("000")) & "') And (EndDate=30000000) And (Active=1)", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											aConceptComponent(N_CENTER_TYPE_ID) = CInt(oRecordset.Fields("CenterTypeID").Value)
										Else
											lErrorNumber = -1
											sErrorDescription = "No existe el tipo de centro de pago de la clave especificada " & aConceptComponent(N_CENTER_TYPE_ID)
										End If
									Else
										lErrorNumber = -1
										sErrorDescription = "No se pudo obtener el tipo de centro de pago de la clave especificada " & aConceptComponent(N_CENTER_TYPE_ID)
									End If
								End If
							Case "StartDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										lStartDateForValueConcept = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lStartDateForValueConcept = CLng(asInputDate(2) & asInputDate(1) & asInputDate(0))
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lStartDateForValueConcept = CLng(asInputDate(2) & asInputDate(0) & asInputDate(1))
								End Select
								If (Err.Number <> 0) Then
									Err.Clear
									sErrorDescription = "Introduzca la fecha de inicio en un formato correcto."
									lErrorNumber = -1
								Else
									aConceptComponent(N_START_DATE_CONCEPT) = CLng(lStartDateForValueConcept)
									If Not VerifyIfUploadMonthDateIsCorrect(aConceptComponent(N_START_DATE_CONCEPT), sErrorDescription) Then
										lErrorNumber = -1
									End If
								End If
							Case "EndDate"
								Select Case sDateFormat
									Case "YYYYMMDD"
										aConceptComponent(N_END_DATE_CONCEPT) = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aConceptComponent(N_END_DATE_CONCEPT) = CLng(asInputDate(2) & asInputDate(1) & asInputDate(0))
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										aConceptComponent(N_END_DATE_CONCEPT) = CLng(asInputDate(2) & asInputDate(0) & asInputDate(1))
								End Select
								If (Err.Number <> 0) Then
									Err.Clear
									sErrorDescription = "Introduzca la fecha de termino en un formato correcto."
									lErrorNumber = -1
								Else
									If Not VerifyIfUploadMonthDateIsCorrect(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), sErrorDescription) Then
										lErrorNumber = -1
									End If
								End If
							Case "IsActive1"
								If lErrorNumber = 0 Then
									If (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 0) Or (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 1) Then
										aConceptComponent(N_IS_ACTIVE1) = CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Else
										lErrorNumber = -1
										sErrorDescription = "Introduzca (<B>1</B>-Si; <B>0</B>-No) como valor para indicar si aplica para guardias"
									End If
								End If
							Case "IsActive2"
								If lErrorNumber = 0 Then
									If (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 0) Or (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 1) Then
										aConceptComponent(N_IS_ACTIVE2) = CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Else
										lErrorNumber = -1
										sErrorDescription = "Introduzca (<B>1</B>-Si; <B>0</B>-No) como valor para indicar si aplica para suplencias"
									End If
								End If
							Case "IsActive3"
								If lErrorNumber = 0 Then
									If (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 0) Or (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 1) Then
										aConceptComponent(N_IS_ACTIVE3) = CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Else
										lErrorNumber = -1
										sErrorDescription = "Introduzca (<B>1</B>-Si; <B>0</B>-No) como valor para indicar si aplica para rezago quirurgico"
									End If
								End If
							Case "IsActive4"
								If lErrorNumber = 0 Then
									If (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 0) Or (CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0)))) = 1) Then
										aConceptComponent(N_IS_ACTIVE4) = CInt(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Else
										lErrorNumber = -1
										sErrorDescription = "Introduzca (<B>1</B>-Si; <B>0</B>-No) como valor para indicar si aplica para PROVAC"
									End If
								End If
							Case Else
						End Select
					Next
					If aConceptComponent(N_END_DATE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000
					If lErrorNumber = 0 Then
						aConceptComponent(N_COMPANY_ID_CONCEPT) = 1
						aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
						aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
						aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
						aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 0
						lErrorNumber = CheckExistencyOfPosition(aConceptComponent, aConceptComponent(S_SHORT_NAME_CONCEPT), sErrorDescription)
					End If
					If lErrorNumber = 0 Then
						If aConceptComponent(N_POSITION_ID_CONCEPT) <> -1 Then
							lErrorNumber = AddPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
						Else
							sErrorDescription = "El puesto no existe en la base de datos."
							lErrorNumber = L_ERR_NO_RECORDS
						End If
					End If
					If lErrorNumber <> 0 Then
						sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex + 1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
					Else
						iTotalRecord = iTotalRecord + 1
					End If
				End If
			Next
		End If
		If iTotalRecord > 0 Then
			Call DisplayErrorMessage("Confirmación", "Han sido registrados " & iTotalRecord & " registros de puestos para guardias y suplencias.")
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If

	UploadPositionsSpecialJourneysLKPFile = lErrorNumber
	Err.Clear
End Function

Function UploadRegisterEmployeesFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         Employees, EmployeesHistoryList, Jobs tables.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadRegisterEmployeesFile"
	Dim oRecordset
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim asInputDate
	Dim iIndex
	Dim jIndex
	Dim iStatusJob
	Dim dConceptAmount
	Dim sAreaShortName
	Dim sComments
	Dim sCURP
	Dim sDate
	Dim sDocumentNumber1
	Dim sDocumentNumber2
	Dim sDocumentNumber3
	Dim sEndDateFormat
	Dim sEndHourFormat
	Dim sEmployeeAddress
	Dim sEmployeeCity
	Dim sEmployeeEmail
	Dim sEmployeeName
	Dim sEmployeeLastName
	Dim sEmployeeLastName2
	Dim sEmployeePhone
	Dim sEmployeeZipCode
	Dim sErrorQueries
	Dim sErrorUpload
	Dim sFields
	Dim sGenderID
	Dim sJourneyShortName
	Dim sOfficePhone
	Dim sOfficeExt
	Dim sPaymentCenterShortName
	Dim sRegisteredErrors
	Dim sQuery	
	Dim sRFC
	Dim sServiceShortName
	Dim sShiftShortName
	Dim sSocialSecurityNumber
	Dim sBirthDateFormat
	Dim sStartDateFormat
	Dim sStartHourFormat
	Dim lCompanyID
	Dim lAreaID
	Dim lEmployeeID
	Dim lJobID
	Dim lEmployeeDate
	Dim lEndDate
	Dim lCountryID
	Dim lBirthDate
	Dim lErrorNumber
	Dim lEmployeeActivityID
	Dim lMaritalStatusID
	Dim lPaymentCenterID
	Dim lStateID
	Dim lShiftID
	Dim lStartHour3
	Dim lEndHour3
	Dim lJourneyID
	Dim lRiskLevel
	Dim lServiceID

	iStatusJob = 0
	dConceptAmount = 0
	sAreaShortName = ""
	sComments = ""
	sCURP = ""
	sDate = ""
	sDocumentNumber1 = ""
	sDocumentNumber2 = ""
	sDocumentNumber3 = ""
	sEmployeeAddress = ""
	sEmployeeCity = ""
	sEmployeeEmail = ""
	sEmployeeName = ""
	sEmployeeLastName = ""
	sEmployeeLastName2 = ""
	sEmployeePhone = ""
	sEmployeeZipCode = ""
	sErrorQueries = ""
	sErrorUpload = ""
	sJourneyShortName = ""
	sOfficePhone = ""
	sOfficeExt = ""
	sPaymentCenterShortName = ""
	sRegisteredErrors = ""
	sQuery = ""
	sRFC = ""
	sServiceShortName = ""
	sShiftShortName = ""
	sSocialSecurityNumber = ""
	lCompanyID = -1
	lAreaID = 0
	lEmployeeID = 0
	lJobID = 0
	lEmployeeDate = 0
	lEndDate = 0
	lCountryID = -1
	lBirthDate = 0
	lErrorNumber = 0
	lMaritalStatusID = 0
	lPaymentCenterID = 0
	lStateID = 0
	lEmployeeActivityID = 0
	lShiftID = 0
	lStartHour3 = 0
	lEndHour3 = 0
	lJourneyID = 0
	lRiskLevel = 0
	lServiceID = 0
	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
		aiFieldsOrder = ""
		aiFieldsOrder = Split(BuildList("-1,", ",", (UBound(asFileRow) + 1)), ",")
		For iIndex = 0 To UBound(asFileRow)
			If IsNull(oRequest("Column" & (iIndex + 1)).Item) Then
			ElseIf StrComp(oRequest("Column" & (iIndex + 1)).Item, "NA", vbBinaryCompare) = 0 Then
			Else
				Select Case lReasonID
					Case 1,2,3,4,5,6,8,10
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "EmployeeID"
								aiFieldsOrder(0) = iIndex & ",EmployeeID"
							Case "OcurredStartDateYYYYMMDD"
								sStartDateFormat = "YYYYMMDD"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateDDMMYYYY"
								sStartDateFormat = "DDMMYYYY"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateMMDDYYYY"
								sStartDateFormat = "MMDDYYYY"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
						End Select
					Case 14
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "EmployeeID"
								aiFieldsOrder(0) = iIndex & ",EmployeeID"
							Case "OcurredStartDateYYYYMMDD"
								sStartDateFormat = "YYYYMMDD"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateDDMMYYYY"
								sStartDateFormat = "DDMMYYYY"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateMMDDYYYY"
								sStartDateFormat = "MMDDYYYY"
								aiFieldsOrder(1) = iIndex & ",EmployeeDate"
							Case "OcurredEndDateYYYYMMDD"
								sEndDateFormat = "YYYYMMDD"
								aiFieldsOrder(2) = iIndex & ",EndDate"
							Case "OcurredEndDateDDMMYYYY"
								sEndDateFormat = "DDMMYYYY"
								aiFieldsOrder(2) = iIndex & ",EndDate"
							Case "OcurredEndDateMMDDYYYY"
								sEndDateFormat = "MMDDYYYY"
								aiFieldsOrder(2) = iIndex & ",EndDate"
							Case "EmployeeName"
								aiFieldsOrder(3) = iIndex & ",EmployeeName"
							Case "EmployeeLastName"
								aiFieldsOrder(4) = iIndex & ",EmployeeLastName"
							Case "EmployeeLastName2"
								aiFieldsOrder(5) = iIndex & ",EmployeeLastName2"
							Case "RFC"
								aiFieldsOrder(6) = iIndex & ",RFC"
							Case "CURP"
								aiFieldsOrder(7) = iIndex & ",CURP"
							Case "SocialSecurityNumber"
								aiFieldsOrder(8) = iIndex & ",SocialSecurityNumber"
							Case "CountryID"
								aiFieldsOrder(9) = iIndex & ",CountryID"
							Case "OcurredBirthDateYYYYMMDD"
								sBirthDateFormat = "YYYYMMDD"
								aiFieldsOrder(10) = iIndex & ",BirthDate"
							Case "OcurredBirthDateDDMMYYYY"
								sBirthDateFormat = "DDMMYYYY"
								aiFieldsOrder(10) = iIndex & ",BirthDate"
							Case "OcurredBirthDateMMDDYYYY"
								sBirthDateFormat = "MMDDYYYY"
								aiFieldsOrder(10) = iIndex & ",BirthDate"
							Case "MaritalStatusID"
								aiFieldsOrder(11) = iIndex & ",MaritalStatusID"
							Case "EmployeeAddress"
								aiFieldsOrder(12) = iIndex & ",EmployeeAddress"
							Case "EmployeeCity"
								aiFieldsOrder(13) = iIndex & ",EmployeeCity"
							Case "EmployeeZipCode"
								aiFieldsOrder(14) = iIndex & ",EmployeeZipCode"
							Case "StateID"
								aiFieldsOrder(15) = iIndex & ",StateID"
							Case "EmployeeEmail"
								aiFieldsOrder(16) = iIndex & ",EmployeeEmail"
							Case "EmployeePhone"
								aiFieldsOrder(17) = iIndex & ",EmployeePhone"
							Case "OfficePhone"
								aiFieldsOrder(18) = iIndex & ",OfficePhone"
							Case "OfficeExt"
								aiFieldsOrder(19) = iIndex & ",OfficeExt"
							Case "DocumentNumber1"
								aiFieldsOrder(20) = iIndex & ",DocumentNumber1"
							Case "DocumentNumber2"
								aiFieldsOrder(21) = iIndex & ",DocumentNumber2"
							Case "DocumentNumber3"
								aiFieldsOrder(22) = iIndex & ",DocumentNumber3"
							Case "EmployeeActivityID"
								aiFieldsOrder(23) = iIndex & ",EmployeeActivityID"
							Case "Comments"
								aiFieldsOrder(24) = iIndex & ",Comments"
							Case "CompanyID"
								aiFieldsOrder(25) = iIndex & ",CompanyID"
							Case "AreaID"
								aiFieldsOrder(26) = iIndex & ",AreaID"
							Case "PaymentCenterID"
								aiFieldsOrder(27) = iIndex & ",PaymentCenterID"
							Case "ServiceID"
								aiFieldsOrder(28) = iIndex & ",ServiceID"
							Case "ConceptAmount"
								aiFieldsOrder(29) = iIndex & ",ConceptAmount"
							Case "JourneyID"
								aiFieldsOrder(30) = iIndex & ",JourneyID"
							Case "ShiftID"
								aiFieldsOrder(31) = iIndex & ",ShiftID"
						End Select
					Case 54
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "JobID"
								aiFieldsOrder(0) = iIndex & ",JobID"
							Case "ServiceID"
								aiFieldsOrder(1) = iIndex & ",ServiceID"
						End Select
					Case Else
						Select Case oRequest("Column" & (iIndex + 1)).Item
							Case "EmployeeID"
								aiFieldsOrder(0) = iIndex & ",EmployeeID"
							Case "JobID"
								aiFieldsOrder(1) = iIndex & ",JobID"
							Case "OcurredStartDateYYYYMMDD"
								sStartDateFormat = "YYYYMMDD"
								aiFieldsOrder(2) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateDDMMYYYY"
								sStartDateFormat = "DDMMYYYY"
								aiFieldsOrder(2) = iIndex & ",EmployeeDate"
							Case "OcurredStartDateMMDDYYYY"
								sStartDateFormat = "MMDDYYYY"
								aiFieldsOrder(2) = iIndex & ",EmployeeDate"
							Case "OcurredEndDateYYYYMMDD"
								sEndDateFormat = "YYYYMMDD"
								aiFieldsOrder(3) = iIndex & ",EndDate"
							Case "OcurredEndDateDDMMYYYY"
								sEndDateFormat = "DDMMYYYY"
								aiFieldsOrder(3) = iIndex & ",EndDate"
							Case "OcurredEndDateMMDDYYYY"
								sEndDateFormat = "MMDDYYYY"
								aiFieldsOrder(3) = iIndex & ",EndDate"
							Case "EmployeeName"
								aiFieldsOrder(4) = iIndex & ",EmployeeName"
							Case "EmployeeLastName"
								aiFieldsOrder(5) = iIndex & ",EmployeeLastName"
							Case "EmployeeLastName2"
								aiFieldsOrder(6) = iIndex & ",EmployeeLastName2"
							Case "RFC"
								aiFieldsOrder(7) = iIndex & ",RFC"
							Case "CURP"
								aiFieldsOrder(8) = iIndex & ",CURP"
							Case "SocialSecurityNumber"
								aiFieldsOrder(9) = iIndex & ",SocialSecurityNumber"
							Case "CountryID"
								aiFieldsOrder(10) = iIndex & ",CountryID"
							Case "OcurredBirthDateYYYYMMDD"
								sBirthDateFormat = "YYYYMMDD"
								aiFieldsOrder(11) = iIndex & ",BirthDate"
							Case "OcurredBirthDateDDMMYYYY"
								sBirthDateFormat = "DDMMYYYY"
								aiFieldsOrder(11) = iIndex & ",BirthDate"
							Case "OcurredBirthDateMMDDYYYY"
								sBirthDateFormat = "MMDDYYYY"
								aiFieldsOrder(11) = iIndex & ",BirthDate"
							Case "MaritalStatusID"
								aiFieldsOrder(12) = iIndex & ",MaritalStatusID"
							Case "EmployeeAddress"
								aiFieldsOrder(13) = iIndex & ",EmployeeAddress"
							Case "EmployeeCity"
								aiFieldsOrder(14) = iIndex & ",EmployeeCity"
							Case "EmployeeZipCode"
								aiFieldsOrder(15) = iIndex & ",EmployeeZipCode"
							Case "StateID"
								aiFieldsOrder(16) = iIndex & ",StateID"
							Case "EmployeeEmail"
								aiFieldsOrder(17) = iIndex & ",EmployeeEmail"
							Case "EmployeePhone"
								aiFieldsOrder(18) = iIndex & ",EmployeePhone"
							Case "OfficePhone"
								aiFieldsOrder(19) = iIndex & ",OfficePhone"
							Case "OfficeExt"
								aiFieldsOrder(20) = iIndex & ",OfficeExt"
							Case "DocumentNumber1"
								aiFieldsOrder(21) = iIndex & ",DocumentNumber1"
							Case "DocumentNumber2"
								aiFieldsOrder(22) = iIndex & ",DocumentNumber2"
							Case "DocumentNumber3"
								aiFieldsOrder(23) = iIndex & ",DocumentNumber3"
							Case "EmployeeActivityID"
								aiFieldsOrder(24) = iIndex & ",EmployeeActivityID"
							Case "ShiftID"
								aiFieldsOrder(25) = iIndex & ",ShiftID"
							Case "Comments"
								aiFieldsOrder(26) = iIndex & ",Comments"
							Case "OcurredStartHour3HHMM"
								sStartHourFormat = "HHMM"
								aiFieldsOrder(27) = iIndex & ",StartHour3"
							Case "OcurredStartHour3HH_MM"
								sStartHourFormat = "HH_MM"
								aiFieldsOrder(27) = iIndex & ",StartHour3"
							Case "OcurredEndHour3HHMM"
								sEndHourFormat = "HHMM"
								aiFieldsOrder(28) = iIndex & ",EndHour3"
							Case "OcurredStartHour3HH_MM"
								sEndHourFormat = "HH_MM"
								aiFieldsOrder(28) = iIndex & ",EndHour3"
							Case "RiskLevel"
								aiFieldsOrder(29) = iIndex & ",RiskLevel"
							Case "ServiceID"
								aiFieldsOrder(30) = iIndex & ",ServiceID"
						End Select
					End Select
			End If
		Next

		sFields = ""
		For iIndex = 0 To UBound(aiFieldsOrder)
			aiFieldsOrder(iIndex) = Split(aiFieldsOrder(iIndex), ",")
			If InStr(1, sFields, aiFieldsOrder(iIndex)(1), vbBinaryCompare) > 0 Then sFields = Replace(sFields, (aiFieldsOrder(iIndex)(1) & ", "), "")
		Next
		If InStr(1, sFields, "EmployeeID") > 0 Then
			lErrorNumber = -1
			sErrorDescription = "La información a registrar no contiene el Número de empleado."
		Else
			sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
			For iIndex = 0 To UBound(asFileContents)
				ReDim Preserve aEmployeeComponent(N_EMPLOYEE_COMPONENT_SIZE)
				If Len(asFileContents(iIndex)) > 0 Then
					asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
					For jIndex = 0 To UBound(aiFieldsOrder)
						If Len(aiFieldsOrder(jIndex)(1)) > 0 Then sQuery = sQuery & aiFieldsOrder(jIndex)(1) & ", "
					Next
					For jIndex = 0 To UBound(aiFieldsOrder)
						Select Case aiFieldsOrder(jIndex)(1)
							Case "ConceptAmount"
								dConceptAmount = CDbl(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "Comments"
								sComments = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeID"
									lEmployeeID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeDate"
								Select Case sStartDateFormat
									Case "YYYYMMDD"
										lEmployeeDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEmployeeDate = CLng(asInputDate(2) & asInputDate(1) & asInputDate(0))
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEmployeeDate = CLng(asInputDate(2) & asInputDate(0) & asInputDate(1))
								End Select
							Case "EndDate"
								Select Case sEndDateFormat
									Case "YYYYMMDD"
										lEndDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEndDate = asInputDate(2) & asInputDate(1) & asInputDate(0)
									Case "MMDDYYYY"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("-")))
										lEndDate = asInputDate(2) & asInputDate(0) & asInputDate(1)
								End Select
							Case "StartHour3"
								Select Case sStartHourFormat
									Case "HHMM"
										lStartHour3 = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "HH_MM"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("_")))
										lStartHour3 = CLng(asInputDate(0)) & CLng(asInputDate(1))
								End Select
							Case "EndHour3"
								Select Case sEndHourFormat
									Case "HHMM"
										lEndHour3 = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "HH_MM"
										asInputDate = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
										asInputDate = Split(asInputDate, Mid(asInputDate, Len("000"), Len("_")))
										lEndHour3 = CLng(asInputDate(0)) & CLng(asInputDate(1))
								End Select
							Case "JobID"
								lJobID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "CompanyID"
								lCompanyID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "AreaID"
								lAreaID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sAreaShortName = Right("00000" & lAreaID,5)
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID From Areas Where AreaCode = '" & sAreaShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lAreaID = CLng(oRecordset.Fields("AreaID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del centro de trabajo no existe.<BR />"
									End If
								End If
							Case "PaymentCenterID"
								lPaymentCenterID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sPaymentCenterShortName = Right("00000" & lPaymentCenterID,5)
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentCenterID From PaymentCenters Where PaymentCenterShortName = '" & sPaymentCenterShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del centro de pago no existe.<BR />"
									End If
								End If
							Case "ServiceID"
								lServiceID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								If Not IsNumeric(lServiceID) Then
									sServiceShortName = Right(("     " & lServiceID), Len("00000"))
								Else
									sServiceShortName = Right(("00000" & lServiceID), Len("00000"))
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceID From Services Where ServiceShortName = '" & sServiceShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lServiceID = CLng(oRecordset.Fields("ServiceID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del servicio no existe.<BR />"
									End If
								End If
							Case "JourneyID"
								lJourneyID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sJourneyShortName = Right("00" & lJourneyID,2)
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyID From Journeys Where JourneyShortName = '" & sJourneyShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lJourneyID = CLng(oRecordset.Fields("JourneyID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del turno no existe.<BR />"
									End If
								End If
							Case "ShiftID"
								lShiftID = asFileRow(CInt(aiFieldsOrder(jIndex)(0)))
								sShiftShortName = Right("0000" & lShiftID,4)
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ShiftID From Shifts Where ShiftShortName = '" & sShiftShortName & "'", "UploadInfoLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lShiftID = CLng(oRecordset.Fields("ShiftID").Value)
									Else
										lErrorNumber = L_ERR_NO_RECORDS
										sErrorUpload = sErrorUpload & "La clave del horario no existe.<BR />"
									End If
								End If
							Case "RiskLevel"
								lRiskLevel = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "BirthDate"
								Select Case sBirthDateFormat
									Case "YYYYMMDD"
										lBirthDate = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
									Case "DDMMYYYY"
										lBirthDate = CLng(asInputDate(2) & asInputDate(1) & asInputDate(0))
									Case "MMDDYYYY"
										lBirthDate = CLng(asInputDate(2) & asInputDate(0) & asInputDate(1))
								End Select
							Case "EmployeeName"
								sEmployeeName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeLastName"
								sEmployeeLastName = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeLastName2"
								sEmployeeLastName2 = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "RFC"
								sRFC = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "CURP"
								sCURP = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeEmail"
								sEmployeeEmail = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "SocialSecurityNumber"
								sSocialSecurityNumber = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "CountryID"
								lCountryID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
								If B_ISSSTE Then
									lCountryID = 0
								End If
							Case "MaritalStatusID"
								lMaritalStatusID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeAddress"
								sEmployeeAddress = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeCity"
								sEmployeeCity = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeZipCode"
								sEmployeeZipCode = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "StateID"
								lStateID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeePhone"
								sEmployeePhone = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "OfficePhone"
								sOfficePhone = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "OfficeExt"
								sOfficeExt = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "DocumentNumber1"
								sDocumentNumber1 = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "DocumentNumber2"
								sDocumentNumber2 = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "DocumentNumber3"
								sDocumentNumber3 = CStr(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
							Case "EmployeeActivityID"
								lEmployeeActivityID = CLng(asFileRow(CInt(aiFieldsOrder(jIndex)(0))))
						End Select
					Next
					Select Case lReasonID
						Case 1, 2, 3, 4, 5, 6, 8, 10, 62, 63, 66
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lEmployeeDate
									aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
									sErrorUpload = "Ocurrió un error al registrar el movimiento al empleado."
									lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
								Else
									sErrorUpload = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
								End If
							End If
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
							End If
						Case 14
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID
								If aEmployeeComponent(N_ID_EMPLOYEE) < 1000000 Then
									If (aEmployeeComponent(N_ID_EMPLOYEE) < 600000) Then
										sErrorUpload = sErrorUpload & "El número de empleado no corresponde a honorarios<BR />"
										lErrorNumber = -1
									End If
								Else
									If (aEmployeeComponent(N_ID_EMPLOYEE) < 1600000) Then
										sErrorUpload = sErrorUpload & "El número de empleado no corresponde a honorarios<BR />"
										lErrorNumber = -1
									End If
								End If
								If lErrorNumber = 0 Then
									lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									sErrorUpload = sErrorUpload & "El número de empleado no existe<BR />"
									lErrorNumber = -1
								End If
								If lErrorNumber = 0 Then
									lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber = 0 Then
										If Len(sRFC) > 0 Then
											aEmployeeComponent(S_RFC_EMPLOYEE) = sRFC
										End If
										If Len(sCURP) > 0 Then
											aEmployeeComponent(S_CURP_EMPLOYEE) = sCURP
										End If
										If lBirthDate <> 0 Then
											aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = lBirthDate
										End If
										If (InStr(1, Left(aEmployeeComponent(S_RFC_EMPLOYEE), Len("000")), Left(aEmployeeComponent(S_CURP_EMPLOYEE), Len("000")), vbBinaryCompare) > 0) Then
											If (InStr(1, Mid(aEmployeeComponent(S_RFC_EMPLOYEE), Len("00000"), Len("000000")), Mid(aEmployeeComponent(S_CURP_EMPLOYEE), Len("00000"), Len("000000")), vbBinaryCompare) > 0) Then
												sGenderID  = Mid(aEmployeeComponent(S_CURP_EMPLOYEE), Len("00000000000"), Len("0"))
												If (InStr(1, sGenderID, "M", vbBinaryCompare) > 0) Then	
													aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 0
												Else
													If (InStr(1, sGenderID, "H", vbBinaryCompare) > 0) Then	
														aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 1
													Else
														sErrorUpload = "La posición 11 del CURP que representa el género del empleado no es válida."
														lErrorNumber = L_ERR_NO_RECORDS
													End If
												End If
												If lErrorNumber = 0 Then
													If (InStr(1, Mid(aEmployeeComponent(S_RFC_EMPLOYEE), Len("00000"), Len("000000")), Mid(CStr(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("000"), Len("000000")), vbBinaryCompare) = 0) Then
														sErrorUpload = "La fecha de nacimiento no coincide con el RFC del empleado."
														lErrorNumber = L_ERR_NO_RECORDS
													End If
												End If
											Else
												sErrorUpload = "El RFC y el CURP del empleado no coinciden"
												lErrorNumber = L_ERR_NO_RECORDS
											End If
										Else
											sErrorUpload = "El RFC y el CURP del empleado no coinciden"
											lErrorNumber = L_ERR_NO_RECORDS
										End If
										Call InitializeJobComponent(oRequest, aJobComponent)
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_ID_EMPLOYEE)
										If lErrorNumber = 0 Then
											If lEmployeeDate <> 0 Then
												aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lEmployeeDate
											End If
											If lEndDate <> 0 Then
												aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)  = lEndDate
											End If
											If Len(sEmployeeName) > 0 Then
												aEmployeeComponent(S_NAME_EMPLOYEE) = sEmployeeName
											End If
											If Len(sEmployeeLastName) > 0 Then
												aEmployeeComponent(S_LAST_NAME_EMPLOYEE) = sEmployeeLastName
											End If
											If Len(sEmployeeLastName2) > 0 Then
												aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = sEmployeeLastName2
											End If
											If Len(sRFC) > 0 Then
												aEmployeeComponent(S_RFC_EMPLOYEE) = sRFC
											End If
											If Len(sCURP) > 0 Then
												aEmployeeComponent(S_CURP_EMPLOYEE) = sCURP
											End If
											If Len(sEmployeeEmail) > 0 Then
												aEmployeeComponent(S_EMAIL_EMPLOYEE) = sEmployeeEmail
											End If
											If Len(sSocialSecurityNumber) > 0 Then
												aEmployeeComponent(S_SSN_EMPLOYEE) = sSocialSecurityNumber
											End If
											If lCountryID <> -1 Then
												aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) = lCountryID
											End If
											If lMaritalStatusID <> 0 Then
												aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) = lMaritalStatusID
											End If
											If Len(sEmployeeAddress) > 0 Then
												aEmployeeComponent(S_ADDRESS_EMPLOYEE) = sEmployeeAddress
											End If
											If Len(sEmployeeCity) > 0 Then
												aEmployeeComponent(S_CITY_EMPLOYEE) = sEmployeeCity
											End If
											If Len(sEmployeeZipCode) > 0 Then
												aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) = sEmployeeZipCode
											End If
											If lStateID <> 0 Then
												aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) = lStateID
											End If
											If Len(sEmployeePhone) > 0 Then
												aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) = sEmployeePhone
											End If
											If Len(sOfficePhone) > 0 Then
												aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) = sOfficePhone
											End If
											If Len(sOfficeExt) > 0 Then
												aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) = sOfficeExt
											End If
											If Len(sDocumentNumber1) > 0 Then
												aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) = sDocumentNumber1
											End If
											If Len(sDocumentNumber2) > 0 Then
												aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) = sDocumentNumber2
											End If
											If Len(sDocumentNumber3) > 0 Then
												aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) = sDocumentNumber3
											End If
											If lEmployeeActivityID <> 0 Then
												aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) = lEmployeeActivityID
											End If
											If lBirthDate <> 0 Then
												aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = lBirthDate
											End If
											If Len(sComments) > 0 Then
												aEmployeeComponent(S_COMMENTS_EMPLOYEE) = sComments
											End If
											If lCompanyID <> -1 Then
												aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = lCompanyID
												aJobComponent(N_COMPANY_ID_JOB) = lCompanyID
											End If
											If lAreaID <> 0 Then
												aEmployeeComponent(N_AREA_ID_EMPLOYEE) = lAreaID
												aJobComponent(N_AREA_ID_JOB) = lAreaID
											End If
											If lPaymentCenterID <> 0 Then
												aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = lPaymentCenterID
												aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
											End If
											If lServiceID <> 0 Then
												aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = lServiceID
												aJobComponent(N_SERVICE_ID_JOB) = lServiceID
											End If
											If lJourneyID <> 0 Then
												aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = lJourneyID
												aJobComponent(N_JOURNEY_ID_JOB) = lJourneyID
											End If
											If lShiftID <> 0 Then
												aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) = lShiftID
												aJobComponent(N_SHIFT_ID_JOB) = lShiftID
											End If
											If dConceptAmount <> 0 Then
												aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = dConceptAmount
											End If
											If lErrorNumber = 0 Then
												aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 3
												aJobComponent(N_POSITION_TYPE_ID_JOB) = 3
												aEmployeeComponent(N_POSITION_ID_EMPLOYEE) = L_HONORARY_POSITION_ID
												aJobComponent(N_POSITION_ID_JOB) = L_HONORARY_POSITION_ID
												lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
												If lErrorNumber <> 0 Then
													sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
												End If
											End If
										Else
											sErrorUpload = "Ocurrió un error al buscar la plaza del empleado."
											lErrorNumber = L_ERR_NO_RECORDS
										End If
									Else
										sErrorUpload = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
										lErrorNumber = L_ERR_NO_RECORDS
									End If
								End If
							End If
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex+1 & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR />"
							End If
						Case 54
							If lErrorNumber = 0 Then
								aJobComponent(N_ID_JOB) = lJobID
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								If Not IsNumeric(lServiceID) Then
									sServiceShortName = Right(("     " & lServiceID), Len("00000"))
								Else
									sServiceShortName = Right(("00000" & lServiceID), Len("00000"))
								End If
								aJobComponent(N_SERVICE_ID_JOB) = CLng(lServiceID)
								aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = False
								aJobComponent(B_IS_DUPLICATED_JOB) = False
								aJobComponent(N_JOB_DATE_JOB) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
								lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							End If
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
							End If
						Case Else
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID
								lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									sErrorUpload = sErrorUpload & "El número de empleado no existe<BR />"
									lErrorNumber = -1
								Else
									lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								End If
								If lErrorNumber = 0 Then
									aEmployeeComponent(N_JOB_ID_EMPLOYEE) = lJobID
									If Len(sRFC) > 0 Then
										aEmployeeComponent(S_RFC_EMPLOYEE) = sRFC
									End If
									If Len(sCURP) > 0 Then
										aEmployeeComponent(S_CURP_EMPLOYEE) = sCURP
									End If
									If lBirthDate <> 0 Then
										aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = lBirthDate
									End If
									If (InStr(1, Left(aEmployeeComponent(S_RFC_EMPLOYEE), Len("000")), Left(aEmployeeComponent(S_CURP_EMPLOYEE), Len("000")), vbBinaryCompare) > 0) Then
										If (InStr(1, Mid(aEmployeeComponent(S_RFC_EMPLOYEE), Len("00000"), Len("000000")), Mid(aEmployeeComponent(S_CURP_EMPLOYEE), Len("00000"), Len("000000")), vbBinaryCompare) > 0) Then
											sGenderID  = Mid(aEmployeeComponent(S_CURP_EMPLOYEE), Len("00000000000"), Len("0"))
											If (InStr(1, sGenderID, "M", vbBinaryCompare) > 0) Then	
												aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 0
											Else
												If (InStr(1, sGenderID, "H", vbBinaryCompare) > 0) Then	
													aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 1
												Else
													sErrorUpload = "La posición 11 del CURP que representa el género del empleado no es válida."
													lErrorNumber = L_ERR_NO_RECORDS
												End If
											End If
											If lErrorNumber = 0 Then
												If (InStr(1, Mid(aEmployeeComponent(S_RFC_EMPLOYEE), Len("00000"), Len("000000")), Mid(CStr(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("000"), Len("000000")), vbBinaryCompare) = 0) Then
													sErrorUpload = "La fecha de nacimiento no coincide con el RFC del empleado."
													lErrorNumber = L_ERR_NO_RECORDS
												End If
											End If
										Else
											sErrorUpload = "El RFC y el CURP del empleado no coinciden"
											lErrorNumber = L_ERR_NO_RECORDS
										End If
									Else
										sErrorUpload = "El RFC y el CURP del empleado no coinciden"
										lErrorNumber = L_ERR_NO_RECORDS
									End If
								End If
								If lErrorNumber = 0 Then
									If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
										lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										If lErrorNumber = 0 Then
											sErrorUpload = "La plaza no tiene el estatus requerido"
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
											If Not oRecordset.EOF Then
												iStatusJob = "," & CStr(oRecordset.Fields("StatusJob2").Value) & ","
											End If
											If lRiskLevel <> 0 Then
												Select Case lRiskLevel
													Case 1,2
														If aJobComponent(N_POSITION_TYPE_ID_JOB) <> 1 Then
															sErrorUpload = "Los riesgos profesionales solo se pueden otorgar a personal de base."
															lErrorNumber = -1
														End If
													Case Else
														sErrorUpload = "La clave del riesgo solo puede ser 1 o 2"
														lErrorNumber = -1
												End Select
											End If
											If (lStartHour3 + lEndHour3) > 0 And (lStartHour3 * lEndHour3) = 0 Then
												sErrorUpload = "El horario de entrada o salida del concepto 07 o 08 son incorrectos."
												lErrorNumber = -1
											Else
												If lStartHour3 <> 0 Then
													If aJobComponent(N_POSITION_TYPE_ID_JOB) <> 1 And aJobComponent(N_POSITION_TYPE_ID_JOB) <> 2 Then
														sErrorUpload = "El concepto 07 y 08 solo pueden otorgarse a personal con puesto de base o confianza."
														lErrorNumber = -1
													End If
													If lErrorNumber = 0 Then
														If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1 Then
															sErrorUpload = "El concepto 07 y 08 no puede otorgarse a funcionarios."
															lErrorNumber = -1
														End If
													End If
												End If
											End If
											If (InStr(1, iStatusJob, "," & aJobComponent(N_STATUS_ID_JOB) & ",", vbBinaryCompare) > 0) Then
												If aJobComponent(N_EMPLOYEE_TYPE_ID_JOB) = aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) Then
													If lEmployeeDate <> 0 Then
														aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lEmployeeDate
													End If
													If lEndDate <> 0 Then
														aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)  = lEndDate
													End If
													If Len(sEmployeeName) > 0 Then
														aEmployeeComponent(S_NAME_EMPLOYEE) = sEmployeeName
													End If
													If Len(sEmployeeLastName) > 0 Then
														aEmployeeComponent(S_LAST_NAME_EMPLOYEE) = sEmployeeLastName
													End If
													If Len(sEmployeeLastName2) > 0 Then
														aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = sEmployeeLastName2
													End If
													If Len(sRFC) > 0 Then
														aEmployeeComponent(S_RFC_EMPLOYEE) = sRFC
													End If
													If Len(sCURP) > 0 Then
														aEmployeeComponent(S_CURP_EMPLOYEE) = sCURP
													End If
													If Len(sEmployeeEmail) > 0 Then
														aEmployeeComponent(S_EMAIL_EMPLOYEE) = sEmployeeEmail
													End If
													If Len(sSocialSecurityNumber) > 0 Then
														aEmployeeComponent(S_SSN_EMPLOYEE) = sSocialSecurityNumber
													End If
													If lCountryID <> -1 Then
														aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) = lCountryID
													End If
													If lMaritalStatusID <> 0 Then
														aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) = lMaritalStatusID
													End If
													If Len(sEmployeeAddress) > 0 Then
														aEmployeeComponent(S_ADDRESS_EMPLOYEE) = sEmployeeAddress
													End If
													If Len(sEmployeeCity) > 0 Then
														aEmployeeComponent(S_CITY_EMPLOYEE) = sEmployeeCity
													End If
													If Len(sEmployeeZipCode) > 0 Then
														aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) = sEmployeeZipCode
													End If
													If lStateID <> 0 Then
														aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) = lStateID
													End If
													If Len(sEmployeePhone) > 0 Then
														aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) = sEmployeePhone
													End If
													If Len(sOfficePhone) > 0 Then
														aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) = sOfficePhone
													End If
													If Len(sOfficeExt) > 0 Then
														aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) = sOfficeExt
													End If
													If Len(sDocumentNumber1) > 0 Then
														aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) = sDocumentNumber1
													End If
													If Len(sDocumentNumber2) > 0 Then
														aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) = sDocumentNumber2
													End If
													If Len(sDocumentNumber3) > 0 Then
														aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) = sDocumentNumber3
													End If
													If lEmployeeActivityID <> 0 Then
														aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) = lEmployeeActivityID
													End If
													If lBirthDate <> 0 Then
														aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = lBirthDate
													End If
													If Len(sComments) > 0 Then
														aEmployeeComponent(S_COMMENTS_EMPLOYEE) = sComments
													End If
													If lServiceID <> 0 Then
														aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = lServiceID
														aJobComponent(N_SERVICE_ID_JOB) = lServiceID
													End If
													If lStartHour3 <> 0 Then
														aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = lStartHour3
													End If
													If lEndHour3 <> 0 Then
														aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = lEndHour3
													End If
													If lRiskLevel <> 0 Then
														aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = lRiskLevel
													End If
													If lErrorNumber = 0 Then
														lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
													End If
												Else
													sErrorUpload = "La plaza que se está asignando no corresponde al mismo tabulador del empleado."
													lErrorNumber = L_ERR_NO_RECORDS
												End If
											Else
												sErrorUpload = "La plaza que se está asignando no tienen estatus requerido."
												lErrorNumber = L_ERR_NO_RECORDS
											End If
										Else
											sErrorUpload = "La plaza no existe."
										End If
									End If
								End If
							End If
							If lErrorNumber <> 0 Then
								sErrorQueries = sErrorQueries & "<B>RENGLÓN " & iIndex & ": </B>" & asFileContents(iIndex) & "<BR /><B>ERROR: </B>" & sErrorUpload & "<BR /><BR />"
							End If
					End Select
				End If
			Next
		End If
		If Len(sErrorQueries) > 0 Then
			lErrorNumber = -1
			sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
		End If
	End If
	UploadRegisterEmployeesFile = lErrorNumber
	Err.Clear
End Function

Function UploadThirdFile(sThirdConcept, sAction, oADODBConnection, sFileName, sOriginalFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         Third table.
'Inputs:  sThirdConcept, sAction, oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadThirdFile"
	Dim oRecordset
	Dim sDate
	Dim aiFieldsOrder
	Dim sFileContents
	Dim asFileContents
	Dim asFileRow
	Dim asInputDate
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim sQuery
	Dim sRow
	Dim sRowOriginal
	Dim sHeaderContents
	Dim lCreditID
	Dim sLoadError
	Dim sFileReportPath
	Dim sFileReportName
	Dim sTarget
	Dim lRFCError
	Dim lEmployeeError
	Dim lEmployeeStatusError
	Dim lConceptError
	Dim lStartDateError
	Dim lDateError
	Dim lPayrollError
	Dim lCreditError
	Dim lAddError
	Dim lTotal
	Dim lTotalSucess
	Dim lTotalError
	Dim lA
	Dim lC
	Dim lB
	Dim sCreditTypeShortName

	Const RFC_ERROR = 1
	Const EMPLOYEE_ERROR = 2
	Const CONCEPT_ERROR = 3
	Const CREDIT_ERROR = 4
	Const START_DATE_ERROR = 5
	Const END_DATE_ERROR = 6

	If Not VerifyExistenceOfFileRegisters(sOriginalFileName, oADODBConnection, sErrorDescription) Then
		sDate = GetSerialNumberForDate("")
		sFileReportPath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & sThirdConcept)
		lErrorNumber = CreateFolder(sFileReportPath, sErrorDescription)
		sFileReportPath = sFileReportPath & "\"
		sFileReportName = sFileReportPath & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "_Rep_" & sThirdConcept & "_" & sDate & ".htm"
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_Third.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) And (Len(asPayrollIDs) > 0) Then
			asPayrollIDs = Left(asPayrollIDs, (Len(asPayrollIDs) - Len(";")))
			asPayrollIDs = Split(asPayrollIDs, ";")
			For iIndex = 0 To UBound(asPayrollIDs)
				asPayrollIDs(iIndex) = Split(asPayrollIDs(iIndex), ",")
				asPayrollIDs(iIndex)(1) = 0
			Next
			sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
			Call GetNameFromTableByShortName(oADODBConnection, "CreditTypes", sThirdConcept, "", "", sCreditTypeShortName, "")
			sHeaderContents = Replace(sHeaderContents, "<THIRD_CONCEPT />", sCreditTypeShortName)
			lErrorNumber = AppendTextToFile(sFileReportName, sHeaderContents, sErrorDescription)
		End If
		lRFCError = 0
		lEmployeeError = 0
		lEmployeeStatusError = 0
		lStartDateError = 0
		lDateError = 0
		lConceptError = 0
		lPayrollError = 0
		lCreditError = 0
		lAddError = 0
		lTotalSucess = 0
		lTotalError = 0
		lA = 0
		lB = 0
		lC = 0
		Call RemoveEmployeeCreditsRejected(oADODBConnection, sOriginalFileName, sErrorDescription)
		Select Case sThirdConcept
			Case THIRD_ISSSTE_CONCEPT
				Dim issste_organismo
				Dim issste_pagaduria
				Dim issste_numero
				Dim issste_fili_issste
				Dim issste_nombre
				Dim issste_adscripcion
				Dim issste_num_empleado
				Dim issste_espacio_1
				Dim issste_tipo_desc
				Dim issste_plazo
				Dim issste_periodo
				Dim aux_issste_periodo
				Dim issste_periodo_venc
				Dim aux_issste_periodo_venc
				Dim issste_concepto
				Dim issste_importe
				Dim issste_importe_decimal
				Dim issste_n_pagaduria
				Dim issste_qna_rep
				Dim issste_fec_rep
				Dim issste_curp
				Dim issste_espacio_2
				Dim issste_tipo_reg
				Dim issste_campo_faltante1

				sFileContents = GetFileContents(sFileName, sErrorDescription)
				If Len(sFileContents) > 0 Then
					asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
					lTotal = UBound(asFileContents) + 1
					For iIndex = 0 To UBound(asFileContents)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
						sRow = asFileContents(iIndex)
						sRowOriginal = sRow
						issste_organismo = Left(sRow, 3)
						sRow = Replace(sRow, issste_organismo, "", 1, 1, vbBinaryCompare)
						issste_pagaduria = Left(sRow, 6)
						sRow = Replace(sRow, issste_pagaduria, "", 1, 1, vbBinaryCompare)
						issste_numero = Left(sRow, 9)
						sRow = Replace(sRow, issste_numero, "", 1, 1, vbBinaryCompare)
						issste_fili_issste = Left(sRow, 13)
						sRow = Replace(sRow, issste_fili_issste, "", 1, 1, vbBinaryCompare)
						issste_nombre = Left(sRow, 40)
						sRow = Replace(sRow, issste_nombre, "", 1, 1, vbBinaryCompare)
						issste_adscripcion = Left(sRow, 10)
						sRow = Replace(sRow, issste_adscripcion, "", 1, 1, vbBinaryCompare)
						issste_num_empleado = Left(sRow, 6)
						sRow = Replace(sRow, issste_num_empleado, "", 1, 1, vbBinaryCompare)
						issste_espacio_1 = Left(sRow, 14)
						sRow = Replace(sRow, issste_espacio_1, "", 1, 1, vbBinaryCompare)
						issste_tipo_desc = Left(sRow, 1)
						sRow = Replace(sRow, issste_tipo_desc, "", 1, 1, vbBinaryCompare)
						issste_plazo = Left(sRow, 3)
						sRow = Replace(sRow, issste_plazo, "", 1, 1, vbBinaryCompare)
						aux_issste_periodo = Left(sRow, 2)
						sRow = Replace(sRow, aux_issste_periodo, "", 1, 1, vbBinaryCompare)
						issste_periodo = Left(sRow, 4)
						sRow = Replace(sRow, issste_periodo, "", 1, 1, vbBinaryCompare)
						'issste_periodo = issste_periodo & aux_issste_periodo
						aux_issste_periodo_venc = Left(sRow, 2)
						sRow = Replace(sRow, aux_issste_periodo_venc, "", 1, 1, vbBinaryCompare)
						issste_periodo_venc = Left(sRow, 4)
						sRow = Replace(sRow, issste_periodo_venc, "", 1, 1, vbBinaryCompare)
						issste_periodo_venc = issste_periodo_venc & aux_issste_periodo_venc
						issste_concepto = Left(sRow, 2)
						sRow = Replace(sRow, issste_concepto, "", 1, 1, vbBinaryCompare)
						issste_importe = Left(sRow, 5)
						sRow = Replace(sRow, issste_importe, "", 1, 1, vbBinaryCompare)
						issste_importe_decimal = Left(sRow, 2)
						sRow = Replace(sRow, issste_importe_decimal, "", 1, 1, vbBinaryCompare)
						issste_importe = issste_importe & "." & issste_importe_decimal
						issste_n_pagaduria = Left(sRow, 6)
						sRow = Replace(sRow, issste_n_pagaduria, "", 1, 1, vbBinaryCompare)
						issste_qna_rep = Left(sRow, 6)
						sRow = Replace(sRow, issste_qna_rep, "", 1, 1, vbBinaryCompare)
						issste_fec_rep = Left(sRow, 8)
						sRow = Replace(sRow, issste_fec_rep, "", 1, 1, vbBinaryCompare)
						issste_curp = Left(sRow, 18)
						sRow = Replace(sRow, issste_curp, "", 1, 1, vbBinaryCompare)
						issste_espacio_2 = Left(sRow, 15)
						sRow = Replace(sRow, issste_espacio_2, "", 1, 1, vbBinaryCompare)
						issste_tipo_reg = Left(sRow, 1)
						sRow = Replace(sRow, issste_tipo_reg, "", 1, 1, vbBinaryCompare)
						issste_campo_faltante1 = Left(sRow, Len(sRow))
						sRow = Replace(sRow, issste_campo_faltante1, "", 1, 1, vbBinaryCompare)
						Select Case issste_concepto
							Case "03"
								issste_concepto = 61
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 61
							Case "08"
								issste_concepto = 82
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 82
						End Select
						Select Case issste_tipo_desc
							Case 1
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1
							Case 2
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3
							Case 3
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 2
						End Select
						aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = issste_plazo
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sOriginalFileName
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						aEmployeeComponent(S_RFC_EMPLOYEE) = Trim(issste_fili_issste)
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(issste_importe)
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = CDbl(issste_importe*issste_plazo)
						lErrorNumber = GetEmployeeNumberFromRFC(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							lRFCError = lRFCError + 1
							sLoadError = sErrorDescription
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
							aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
							Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
						Else
							lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								lRFCError = lRFCError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
							End If
						End If
						If lErrorNumber = 0 Then
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = -1
								lEmployeeStatusError = lEmployeeStatusError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
							Else
								lErrorNumber = GetCreditsStartDate(issste_periodo, aux_issste_periodo, aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE))
								If lErrorNumber <> 0 Then
									lErrorNumber = -1
									lStartDateError = lStartDateError + 1
									sLoadError = "No se pudo obtener una fecha de inicio para el año " & issste_periodo & " del periodo " & aux_issste_periodo
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, START_DATE_ERROR, sErrorDescription)
								Else
									'aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) = CLng(oRecordset.Fields("PayrollID").Value)
									lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										lErrorNumber = -1
										lDateError	lDateError 	= lDateError + 1
										sLoadError = "Error al obtener la fecha de término del tercero."
										aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
										aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
										Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, END_DATE_ERROR, sErrorDescription)
									Else
										If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If (aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1) And (VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 0, sErrorDescription)) Then
											lErrorNumber = -1
											lCreditError = lCreditError + 1
											'sLoadError = "Ya existe un registro del mismo tipo para validación"
											sLoadError = sErrorDescription
										Else
											aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
											lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												lAddError = lAddError + 1
												sLoadError = "No se pudo agregar el crédito del empleado"
											Else
												Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
													Case 1
														lA = lA + 1
													Case 3
														lB = lB + 1
													Case 2
														lC = lC + 1
												End select
											End If
										End If
									End If
								End If
							End If
						End If
						sOutputText = ""
						If lErrorNumber = 0 Then
							lTotalSucess = lTotalSucess + 1
							sOutputText = sOutputText & "<TR>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("OK")) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
							sOutputText = sOutputText & "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						Else
							lTotalError = lTotalError + 1
							sOutputText = sOutputText & "<TR>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("Error")) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(sLoadError)) & "</FONT></TD>"
							sOutputText = sOutputText &  "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						End If
					Next
					sOutputText = ""
					sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de número de empleado:&nbsp;" & CleanStringForHTML(CStr(lRFCError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de estatus del empleado:&nbsp;" & CleanStringForHTML(CStr(lEmployeeStatusError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha inicial:&nbsp;" & CleanStringForHTML(CStr(lStartDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha final:&nbsp;" & CleanStringForHTML(CStr(lDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de clave de concepto:&nbsp;" & CleanStringForHTML(CStr(lConceptError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de registro duplicado:&nbsp;" & CleanStringForHTML(CStr(lCreditError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores al agregar el crédito:&nbsp;" & CleanStringForHTML(CStr(lAddError)) & "</B></FONT></TD></TR>"
					'sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros :&nbsp;" & CleanStringForHTML(CStr(lTotal)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros correctos:&nbsp;" & CleanStringForHTML(CStr(lTotalSucess)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Altas:&nbsp;" & CleanStringForHTML(CStr(lA)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Cambios:&nbsp;" & CleanStringForHTML(CStr(lC)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Bajas:&nbsp;" & CleanStringForHTML(CStr(lB)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros rechazados:&nbsp;" & CleanStringForHTML(CStr(lTotalError)) & "</B></FONT></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
				End If
				sTarget = Replace(sFileReportName, SYSTEM_PHYSICAL_PATH, "", 1, 1, vbBinaryCompare)
				sTarget = "<A HREF="" " & sTarget & """ target=""_blank"">Ver el informe completo de carga</A>"
				Call DisplayErrorMessage("La información ha sido procesada", sTarget)
			Case THIRD_FOVISSSTE_CONCEPT_86, THIRD_FOVISSSTE_CONCEPT_56
				Dim fovissste_filiacion
				Dim fovissste_nombre
				Dim fovissste_camp1
				Dim fovissste_cve_mov
				Dim fovissste_importe
				Dim fovissste_importe_decimal
				Dim fovissste_concepto
				Dim fovissste_quincena
				Dim fovissste_percen
				Dim fovissste_errores
				Dim fovissste_estatus
				Dim fovissste_campo_faltante

				sFileContents = GetFileContents(sFileName, sErrorDescription)
				If Len(sFileContents) > 0 Then
					asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
					lTotal = UBound(asFileContents) + 1
					For iIndex = 0 To UBound(asFileContents)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
						sRow = asFileContents(iIndex)
						sRowOriginal = sRow
						fovissste_filiacion = Left(sRow, 13)
						sRow = Replace(sRow, fovissste_filiacion, "", 1, 1, vbBinaryCompare)
						fovissste_nombre = Left(sRow, 30)
						sRow = Replace(sRow, fovissste_nombre, "", 1, 1, vbBinaryCompare)
						fovissste_camp1 = Left(sRow, 37)
						sRow = Replace(sRow, fovissste_camp1, "", 1, 1, vbBinaryCompare)
						fovissste_cve_mov = Left(sRow, 1)
						sRow = Replace(sRow, fovissste_cve_mov, "", 1, 1, vbBinaryCompare)
						fovissste_importe = Left(sRow, 6)
						sRow = Replace(sRow, fovissste_importe, "", 1, 1, vbBinaryCompare)
						fovissste_importe_decimal = Left(sRow, 2)
						sRow = Replace(sRow, fovissste_importe_decimal, "", 1, 1, vbBinaryCompare)
						fovissste_importe = fovissste_importe & "." & fovissste_importe_decimal
						fovissste_concepto = Left(sRow, 3)
						sRow = Replace(sRow, fovissste_concepto, "", 1, 1, vbBinaryCompare)
						If sThirdConcept = THIRD_FOVISSSTE_CONCEPT_56 Then
							fovissste_quincena = Left(sRow, 2)
							sRow = Replace(sRow, fovissste_quincena, "", 1, 1, vbBinaryCompare)
							fovissste_percen = Left(sRow, 6)
							sRow = Replace(sRow, fovissste_percen, "", 1, 1, vbBinaryCompare)
						Else
							fovissste_quincena = Left(sRow, 6)
							sRow = Replace(sRow, fovissste_quincena, "", 1, 1, vbBinaryCompare)
							fovissste_percen = Left(sRow, 2)
							sRow = Replace(sRow, fovissste_percen, "", 1, 1, vbBinaryCompare)
						End If
						If InStr(1, fovissste_concepto, "56L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "56"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 58
						ElseIf InStr(1, fovissste_concepto, "64L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "62"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 64
						ElseIf InStr(1, fovissste_concepto, "65L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "86"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 83
						End If
						Select Case UCase(fovissste_cve_mov)
							Case "A"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1
							Case "B"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3
							Case "C"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 2
						End Select
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sOriginalFileName
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						aEmployeeComponent(S_RFC_EMPLOYEE) = Trim(fovissste_filiacion)
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(fovissste_importe)
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = 0
						lErrorNumber = GetEmployeeNumberFromRFC(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							lRFCError = lRFCError + 1
							sLoadError = sErrorDescription
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
							aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
							Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
						Else
							lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								lRFCError = lRFCError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
							End If
						End If
						If lErrorNumber = 0 Then
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = -1
								lEmployeeStatusError = lEmployeeStatusError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
							Else
								If sThirdConcept = THIRD_FOVISSSTE_CONCEPT_56 Then
									'lErrorNumber = GetCreditsStartDate(Left(fovissste_quincena, Len("0000")), Right(fovissste_quincena, Len("00")), aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE))
									aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = CInt(fovissste_quincena)
									aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = GetPayrollStartDate(GetSerialNumberForDate(""))
									sLoadError = "No se pudo obtener una fecha de inicio para el número de periodos indicado: " & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE)
								Else
									lErrorNumber = GetCreditsStartDate(Left(fovissste_quincena, Len("0000")), Right(fovissste_quincena, Len("00")), aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE))
									sLoadError = "No se pudo obtener una fecha de inicio para el año " & Left(fovissste_quincena, Len("0000")) & " del periodo " & Right(fovissste_quincena, Len("00"))
								End If
								If lErrorNumber <> 0 Then
									lErrorNumber = -1
									lStartDateError = lStartDateError + 1
									'sLoadError = "No se pudo obtener una fecha de inicio para el año " & issste_periodo & " del periodo " & aux_issste_periodo
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, START_DATE_ERROR, sErrorDescription)
								Else
									lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										lErrorNumber = -1
										lDateError = lDateError + 1
										sLoadError = "Error al obtener la fecha de término del tercero."
										aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
										aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
										Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, END_DATE_ERROR, sErrorDescription)
									Else
										If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If (aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1) And (VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 0, sErrorDescription)) Then
											lErrorNumber = -1
											lCreditError = lCreditError + 1
											'sLoadError = "Ya existe un registro del mismo tipo para validación"
											sLoadError = sErrorDescription
										Else
											aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
											lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												lAddError = lAddError + 1
												sLoadError = "No se pudo agregar el crédito del empleado debido a que " & sErrorDescription
											Else
												Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
													Case 1
														lA = lA + 1
													Case 3
														lB = lB + 1
													Case 2
														lC = lC + 1
												End select
											End If
										End If
									End If
								End If
							End If
						End If
						sOutputText = ""
						If lErrorNumber = 0 Then
							lTotalSucess 	= lTotalSucess + 1
							sOutputText = sOutputText & "<TR>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("OK")) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
							sOutputText = sOutputText & "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						Else
							lTotalError 	= lTotalError + 1
							sOutputText = sOutputText & "<TR>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("Error")) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(sLoadError)) & "</FONT></TD>"
							sOutputText = sOutputText &  "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						End If
					Next
					sOutputText = ""
					sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de número de empleado:&nbsp;" & CleanStringForHTML(CStr(lRFCError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de estatus del empleado:&nbsp;" & CleanStringForHTML(CStr(lEmployeeStatusError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha inicial:&nbsp;" & CleanStringForHTML(CStr(lStartDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha final:&nbsp;" & CleanStringForHTML(CStr(lDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de Clave de concepto:&nbsp;" & CleanStringForHTML(CStr(lConceptError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de registro duplicado:&nbsp;" & CleanStringForHTML(CStr(lCreditError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores al agregar el crédito:&nbsp;" & CleanStringForHTML(CStr(lAddError)) & "</B></FONT></TD></TR>"

					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros :&nbsp;" & CleanStringForHTML(CStr(lTotal)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros correctos:&nbsp;" & CleanStringForHTML(CStr(lTotalSucess)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Altas:&nbsp;" & CleanStringForHTML(CStr(lA)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Cambios:&nbsp;" & CleanStringForHTML(CStr(lC)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Bajas:&nbsp;" & CleanStringForHTML(CStr(lB)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros rechazados:&nbsp;" & CleanStringForHTML(CStr(lTotalError)) & "</B></FONT></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
				End If
				sTarget = Replace(sFileReportName, SYSTEM_PHYSICAL_PATH, "", 1, 1, vbBinaryCompare)
				sTarget = "<A HREF="" " & sTarget & """ target=""_blank"">Ver informe completo de carga</A>"
				Call DisplayErrorMessage("La información ha sido procesada", sTarget)
			Case THIRD_FOVISSSTE_CONCEPT_62
				Dim fovissste_filiacion_62
				Dim fovissste_nombre_62
				Dim fovissste_camp1_62
				Dim fovissste_cve_mov_62
				Dim fovissste_importe_62
				Dim fovissste_importe_decimal_62
				Dim fovissste_concepto_62
				Dim fovissste_quincena_62
				Dim fovissste_percen_62
				Dim fovissste_errores_62
				Dim fovissste_estatus_62
				Dim fovissste_campo_faltante_62

				sFileContents = GetFileContents(sFileName, sErrorDescription)
				If Len(sFileContents) > 0 Then
					asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
					lTotal = UBound(asFileContents) + 1
					For iIndex = 0 To UBound(asFileContents)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
						sRow = asFileContents(iIndex)
						sRowOriginal = sRow
						fovissste_filiacion_62 = Left(sRow, 13)
						sRow = Replace(sRow, fovissste_filiacion_62, "", 1, 1, vbBinaryCompare)
						fovissste_nombre_62 = Left(sRow, 30)
						sRow = Replace(sRow, fovissste_nombre_62, "", 1, 1, vbBinaryCompare)
						fovissste_camp1_62 = Left(sRow, 37)
						sRow = Replace(sRow, fovissste_camp1_62, "", 1, 1, vbBinaryCompare)
						fovissste_cve_mov_62 = Left(sRow, 1)
						sRow = Replace(sRow, fovissste_cve_mov_62, "", 1, 1, vbBinaryCompare)
						fovissste_importe_62 = Left(sRow, 5)
						sRow = Replace(sRow, fovissste_importe_62, "", 1, 1, vbBinaryCompare)
						fovissste_importe_decimal_62 = Left(sRow, 3)
						sRow = Replace(sRow, fovissste_importe_decimal_62, "", 1, 1, vbBinaryCompare)
						fovissste_importe_62 = fovissste_importe_62 & "." & fovissste_importe_decimal_62
						fovissste_concepto_62 = Left(sRow, 3)
						sRow = Replace(sRow, fovissste_concepto_62, "", 1, 1, vbBinaryCompare)
						fovissste_quincena_62 = Left(sRow, 6)
						sRow = Replace(sRow, fovissste_quincena_62, "", 1, 1, vbBinaryCompare)
						fovissste_percen_62 = Left(sRow, 2)
						sRow = Replace(sRow, fovissste_percen_62, "", 1, 1, vbBinaryCompare)
						If InStr(1, fovissste_concepto_62, "56L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "56"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 58
						ElseIf InStr(1, fovissste_concepto_62, "64L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "62"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 64
						ElseIf InStr(1, fovissste_concepto_62, "65L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "86"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 83
						End If
						Select Case UCase(fovissste_cve_mov_62)
							Case "A"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1
							Case "B"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3
							Case "C"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 2
						End Select
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sOriginalFileName
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						aEmployeeComponent(S_RFC_EMPLOYEE) = Trim(fovissste_filiacion_62)
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CInt(fovissste_percen_62)
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = 0
						lErrorNumber = GetEmployeeNumberFromRFC(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							lRFCError = lRFCError + 1
							sLoadError = sErrorDescription
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
							aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
							Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
						Else
							lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								lRFCError = lRFCError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
							End If
						End If
						If lErrorNumber = 0 Then
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = -1
								lEmployeeStatusError = lEmployeeStatusError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
							Else
								lErrorNumber = GetCreditsStartDate(Left(fovissste_quincena_62, Len("0000")), Right(fovissste_quincena_62, Len("00")), aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE))
								If lErrorNumber <> 0 Then
									lErrorNumber = -1
									lStartDateError = lStartDateError + 1
									sLoadError = "No se pudo obtener una fecha de inicio para el año " & issste_periodo_62 & " del periodo " & aux_issste_periodo_62
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, START_DATE_ERROR, sErrorDescription)
								Else
									lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										lErrorNumber = -1
										lDateError = lDateError + 1
										sLoadError = "Error al obtener la fecha de término del tercero."
										aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
										aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
										Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, END_DATE_ERROR, sErrorDescription)
									Else
										If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If (aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1) And (VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 0, sErrorDescription)) Then
											lErrorNumber = -1
											lCreditError = lCreditError + 1
											'sLoadError = "Ya existe un registro del mismo tipo para validación"
											sLoadError = sErrorDescription
										Else
											aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
											lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												lAddError = lAddError + 1
												sLoadError = "No se pudo agregar el crédito del empleado debido a que " & sErrorDescription
											Else
												Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
													Case 1
														lA = lA + 1
													Case 3
														lB = lB + 1
													Case 2
														lC = lC + 1
												End select
											End If
										End If
									End If
								End If
							End If
						End If
						sOutputText = ""
						If lErrorNumber = 0 Then
							lTotalSucess 	= lTotalSucess + 1
							sOutputText = sOutputText & "<TR>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("OK")) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
							sOutputText = sOutputText & "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						Else
							lTotalError 	= lTotalError + 1
							sOutputText = sOutputText & "<TR>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("Error")) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(sLoadError)) & "</FONT></TD>"
							sOutputText = sOutputText &  "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						End If
					Next
					sOutputText = ""
					sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de número de empleado:&nbsp;" & CleanStringForHTML(CStr(lRFCError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de estatus del empleado:&nbsp;" & CleanStringForHTML(CStr(lEmployeeStatusError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha inicial:&nbsp;" & CleanStringForHTML(CStr(lStartDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha final:&nbsp;" & CleanStringForHTML(CStr(lDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de Clave de concepto:&nbsp;" & CleanStringForHTML(CStr(lConceptError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de registro duplicado:&nbsp;" & CleanStringForHTML(CStr(lCreditError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores al agregar el crédito:&nbsp;" & CleanStringForHTML(CStr(lAddError)) & "</B></FONT></TD></TR>"

					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros :&nbsp;" & CleanStringForHTML(CStr(lTotal)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros correctos:&nbsp;" & CleanStringForHTML(CStr(lTotalSucess)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Altas:&nbsp;" & CleanStringForHTML(CStr(lA)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Cambios:&nbsp;" & CleanStringForHTML(CStr(lC)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Bajas:&nbsp;" & CleanStringForHTML(CStr(lB)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros rechazados:&nbsp;" & CleanStringForHTML(CStr(lTotalError)) & "</B></FONT></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
				End If
				sTarget = Replace(sFileReportName, SYSTEM_PHYSICAL_PATH, "", 1, 1, vbBinaryCompare)
				sTarget = "<A HREF="" " & sTarget & """ target=""_blank"">Ver informe completo de carga</A>"
				Call DisplayErrorMessage("La información ha sido procesada", sTarget)
			Case THIRD_FOVISSSTE_CONCEPT_NF
				Dim fovissste_filiacion_NF
				Dim fovissste_nombre_NF
				Dim fovissste_camp1_NF
				Dim fovissste_cve_mov_NF
				Dim fovissste_importe_NF
				Dim fovissste_importe_decimal_NF
				Dim fovissste_concepto_NF
				Dim fovissste_quincena_NF
				Dim fovissste_percen_NF
				Dim fovissste_errores_NF
				Dim fovissste_estatus_NF
				Dim fovissste_campo_faltante_NF

				sFileContents = GetFileContents(sFileName, sErrorDescription)
				If Len(sFileContents) > 0 Then
					asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
					lTotal = UBound(asFileContents) + 1
					For iIndex = 0 To UBound(asFileContents)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
						sRow = asFileContents(iIndex)
						sRowOriginal = sRow
						fovissste_filiacion_NF = Left(sRow, 13)
						sRow = Replace(sRow, fovissste_filiacion_NF, "", 1, 1, vbBinaryCompare)
						fovissste_nombre_NF = Left(sRow, 30)
						sRow = Replace(sRow, fovissste_nombre_NF, "", 1, 1, vbBinaryCompare)
						fovissste_camp1_NF = Left(sRow, 33)
						sRow = Replace(sRow, fovissste_camp1_NF, "", 1, 1, vbBinaryCompare)
						fovissste_cve_mov_NF = Left(sRow, 1)
						sRow = Replace(sRow, fovissste_cve_mov_NF, "", 1, 1, vbBinaryCompare)
						fovissste_concepto_NF = Left(sRow, 3)
						sRow = Replace(sRow, fovissste_concepto_NF, "", 1, 1, vbBinaryCompare)

						fovissste_importe_NF = Left(sRow, 6)
						sRow = Replace(sRow, fovissste_importe_NF, "", 1, 1, vbBinaryCompare)
						fovissste_importe_decimal_NF = Left(sRow, 2)
						sRow = Replace(sRow, fovissste_importe_decimal_NF, "", 1, 1, vbBinaryCompare)
						fovissste_importe_NF = fovissste_importe_NF & "." & fovissste_importe_decimal_NF

						fovissste_quincena_NF = Left(sRow, 6)
						sRow = Replace(sRow, fovissste_quincena_NF, "", 1, 1, vbBinaryCompare)
						fovissste_percen_NF = Left(sRow, 3)
						sRow = Replace(sRow, fovissste_percen_NF, "", 1, 1, vbBinaryCompare)
						If InStr(1, fovissste_concepto_NF, "56L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "56"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 58
						ElseIf InStr(1, fovissste_concepto_NF, "64L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "62"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 64
						ElseIf InStr(1, fovissste_concepto_NF, "65L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "86"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 83
						ElseIf InStr(1, fovissste_concepto_NF, "55L", vbBinaryCompare) > 0 Then
							fovissste_concepto = "NF"
							fovissste_errores= "     "
							fovissste_estatus= "  "
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 58
						End If
						Select Case UCase(fovissste_cve_mov_NF)
							Case "A"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1
							Case "B"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3
							Case "C"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 2
						End Select
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sOriginalFileName
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						aEmployeeComponent(S_RFC_EMPLOYEE) = Trim(fovissste_filiacion_NF)
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(fovissste_importe_NF)
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = 0
						lErrorNumber = GetEmployeeNumberFromRFC(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							lRFCError = lRFCError + 1
							sLoadError = sErrorDescription
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
							aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
							Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
						Else
							lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								lRFCError = lRFCError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
							End If
						End If
						If lErrorNumber = 0 Then
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = -1
								lEmployeeStatusError = lEmployeeStatusError + 1
								sLoadError = sErrorDescription
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
								aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
								Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
							Else
								'lErrorNumber = GetCreditsStartDate(Left(fovissste_quincena_NF, Len("0000")), Right(fovissste_quincena_NF, Len("00")), aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE))
								aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = CInt(Right(fovissste_quincena_NF, Len("00")))
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = GetPayrollStartDate(GetSerialNumberForDate(""))
								If lErrorNumber <> 0 Then
									lErrorNumber = -1
									lStartDateError = lStartDateError + 1
									sLoadError = "No se pudo obtener una fecha de inicio para el número de periodos indicado: " & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE)
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, START_DATE_ERROR, sErrorDescription)
								Else
									lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										lErrorNumber = -1
										lDateError = lDateError + 1
										sLoadError = "Error al obtener la fecha de término del tercero."
										aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
										aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
										Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, END_DATE_ERROR, sErrorDescription)
									Else
										If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
										If (aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1) And (VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 0, sErrorDescription)) Then
											lErrorNumber = -1
											lCreditError = lCreditError + 1
											'sLoadError = "Ya existe un registro del mismo tipo para validación"
											sLoadError = sErrorDescription
										Else
											aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
											lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												lAddError = lAddError + 1
												sLoadError = "No se pudo agregar el crédito del empleado debido a que " & sErrorDescription
											Else
												Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
													Case 1
														lA = lA + 1
													Case 3
														lB = lB + 1
													Case 2
														lC = lC + 1
												End select
											End If
										End If
									End If
								End If
							End If
						End If
						sOutputText = ""
						If lErrorNumber = 0 Then
							lTotalSucess 	= lTotalSucess + 1
							sOutputText = sOutputText & "<TR>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("OK")) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
							sOutputText = sOutputText & "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						Else
							lTotalError 	= lTotalError + 1
							sOutputText = sOutputText & "<TR>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("Error")) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(sLoadError)) & "</FONT></TD>"
							sOutputText = sOutputText &  "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						End If
					Next
					sOutputText = ""
					sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de número de empleado:&nbsp;" & CleanStringForHTML(CStr(lRFCError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de estatus del empleado:&nbsp;" & CleanStringForHTML(CStr(lEmployeeStatusError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha inicial:&nbsp;" & CleanStringForHTML(CStr(lStartDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha final:&nbsp;" & CleanStringForHTML(CStr(lDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de Clave de concepto:&nbsp;" & CleanStringForHTML(CStr(lConceptError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de registro duplicado:&nbsp;" & CleanStringForHTML(CStr(lCreditError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores al agregar el crédito:&nbsp;" & CleanStringForHTML(CStr(lAddError)) & "</B></FONT></TD></TR>"

					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros :&nbsp;" & CleanStringForHTML(CStr(lTotal)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros correctos:&nbsp;" & CleanStringForHTML(CStr(lTotalSucess)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Altas:&nbsp;" & CleanStringForHTML(CStr(lA)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Cambios:&nbsp;" & CleanStringForHTML(CStr(lC)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Bajas:&nbsp;" & CleanStringForHTML(CStr(lB)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros rechazados:&nbsp;" & CleanStringForHTML(CStr(lTotalError)) & "</B></FONT></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
				End If
				sTarget = Replace(sFileReportName, SYSTEM_PHYSICAL_PATH, "", 1, 1, vbBinaryCompare)
				sTarget = "<A HREF="" " & sTarget & """ target=""_blank"">Ver informe completo de carga</A>"
				Call DisplayErrorMessage("La información ha sido procesada", sTarget)
			Case Else
				Dim third_Empresa
				Dim third_NumeroDeEmpleado
				Dim third_Nombre
				Dim third_Rfc
				Dim third_Curp
				Dim third_Tipo
				Dim third_Plazo
				Dim third_PeriodoInicio_Dia
				Dim third_PeriodoInicio_Mes
				Dim third_PeriodoInicio_Anio
				Dim third_PeriodoTermino
				Dim third_Importe
				Dim aux_third_Importe
				Dim third_QuincenaProceso
				Dim third_Concepto
				Dim sOutputText
				Dim third_PayrollDate

				sFileContents = GetFileContents(sFileName, sErrorDescription)
				If Len(sFileContents) > 0 Then
					asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
					lTotal = UBound(asFileContents) + 1
					For iIndex = 0 To UBound(asFileContents)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
						sRow = asFileContents(iIndex)
						sRowOriginal = sRow
						third_Empresa = Left(sRow, 1)
						sRow = Replace(sRow, third_Empresa, "", 1, 1, vbBinaryCompare)
						third_NumeroDeEmpleado = Left(sRow, 6)
						sRow = Replace(sRow, third_NumeroDeEmpleado, "", 1, 1, vbBinaryCompare)
						third_Nombre = Left(sRow, 50)
						sRow = Replace(sRow, third_Nombre, "", 1, 1, vbBinaryCompare)
						third_Rfc = Left(sRow, 13)
						sRow = Replace(sRow, third_Rfc, "", 1, 1, vbBinaryCompare)
						third_Curp = Left(sRow, 18)
						sRow = Replace(sRow, third_Curp, "", 1, 1, vbBinaryCompare)
						third_Tipo = Left(sRow, 1)
						sRow = Replace(sRow, third_Tipo, "", 1, 1, vbBinaryCompare)
						third_Plazo = Left(sRow, 3)
						sRow = Replace(sRow, third_Plazo, "", 1, 1, vbBinaryCompare)
						'third_PeriodoInicio = Left(sRow, 6)
						'sRow = Replace(sRow, third_PeriodoInicio, "", 1, 1, vbBinaryCompare)
						third_PeriodoInicio_Dia = Left(sRow, 2)
						sRow = Replace(sRow, third_PeriodoInicio_Dia, "", 1, 1, vbBinaryCompare)
						third_PeriodoInicio_Mes = Left(sRow, 2)
						sRow = Replace(sRow, third_PeriodoInicio_Mes, "", 1, 1, vbBinaryCompare)
						third_PeriodoInicio_Anio = Left(sRow, 2)
						sRow = Replace(sRow, third_PeriodoInicio_Anio, "", 1, 1, vbBinaryCompare)
						third_PeriodoTermino = Left(sRow, 6)
						sRow = Replace(sRow, third_PeriodoTermino, "", 1, 1, vbBinaryCompare)
						third_Importe = Left(sRow, 17)
						sRow = Replace(sRow, third_Importe, "", 1, 1, vbBinaryCompare)
						aux_third_Importe = Left(sRow, 2)
						sRow = Replace(sRow, aux_third_Importe, "", 1, 1, vbBinaryCompare)
						third_Importe = third_Importe & "." & aux_third_Importe
						third_QuincenaProceso = Left(sRow, 2)
						sRow = Replace(sRow, third_QuincenaProceso, "", 1, 1, vbBinaryCompare)
						third_Concepto = Left(sRow, 2)
						sRow = Replace(sRow, third_Concepto, "", 1, 1, vbBinaryCompare)
						Select Case third_Concepto
							Case "72"
								third_Concepto = 75
							Case Else
						End Select
						aEmployeeComponent(N_ID_EMPLOYEE) = CLng(third_NumeroDeEmpleado)
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(third_Importe)
						aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = CLng(third_Plazo)
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = CDbl(third_Importe*third_Plazo)
						Select Case third_Tipo
							Case "A"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1
							Case "B"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 3
								aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = 0
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 0
							Case "C"
								aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 2
						End Select
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = sOriginalFileName
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						aEmployeeComponent(S_RFC_EMPLOYEE) = Trim(third_Rfc)
						third_PayrollDate = CLng("20" & Left(third_PeriodoInicio, Len("00")) & Left(third_QuincenaProceso, Len("00")))
						aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng("20" & third_PeriodoInicio_Anio & third_PeriodoInicio_Mes & third_PeriodoInicio_Dia)
						If  Not VerifyIfUploadMonthDateIsCorrect(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
							lErrorNumber = -1
							lStartDateError = lStartDateError + 1
							sLoadError = sErrorDescription
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
							aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
							Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, START_DATE_ERROR, sErrorDescription)
						Else
							lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								lErrorNumber = GetEmployeeNumberFromRFC(oRequest, oADODBConnection, 0, aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									lRFCError = lRFCError + 1
									sLoadError = "No existe empleado con el número indicado. "  & sErrorDescription
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, EMPLOYEE_ERROR, sErrorDescription)
								End If
							End If
							If lErrorNumber = 0 Then
								If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
									lErrorNumber = -1
									lEmployeeStatusError = lEmployeeStatusError + 1
									sLoadError = sErrorDescription
									aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
									aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
									Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
								Else
									sQuery = "Select * from CreditTypes Where (CreditTypeShortName = '" & third_Concepto & "')"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber <> 0 Then
										lErrorNumber = -1
										lConceptError = lConceptError + 1
										sLoadError = "Error al obtener la clave del concepto: " & third_Concepto
										aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
										aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
										Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
									Else
										If oRecordset.EOF Then
											lErrorNumber = -1
											lConceptError = lConceptError + 1
											sLoadError = "No existe la clave del concepto indicado: " & third_Concepto
											aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
											aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
											Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
										Else
											aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditTypeID").Value)
											lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												lErrorNumber = -1
												lDateError = lDateError + 1
												sLoadError = "Error al obtener la fecha de término del tercero."
												aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
												aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
												Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, END_DATE_ERROR, sErrorDescription)
											Else
												If (aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 1) And (VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 0, sErrorDescription)) Then
													lErrorNumber = -1
													lCreditError = lCreditError + 1
													sLoadError = "Ya existe un registro del mismo tipo en proceso para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
													aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
													aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
													Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CONCEPT_ERROR, sErrorDescription)
												Else
													aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = ""
													lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
													If lErrorNumber <> 0 Then
														lAddError = lAddError + 1
														sLoadError = "No se pudo agregar el crédito del empleado"
														aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) = iIndex
														aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS) = sLoadError
														Call AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, CREDIT_ERROR, sErrorDescription)
													Else
														Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
															Case 1
																lA = lA + 1
															Case 3
																lB = lB + 1
															Case 2
																lC = lC + 1
														End select
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
						sOutputText = ""
						If lErrorNumber = 0 Then
							lTotalSucess = lTotalSucess + 1
							sOutputText = sOutputText & "<TR>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("OK")) & "</FONT></TD>"
								sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
									Case 1
										sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro de alta para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
									Case 3
										sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro de baja para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
									Case 2
										sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("El registro de cambio para el número de empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " fue realizado con éxito.")) & "</FONT></TD>"
								End Select
							sOutputText = sOutputText & "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						Else
							lTotalError = lTotalError + 1
							sOutputText = sOutputText & "<TR>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(iIndex+1)) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr("Error")) & "</FONT></TD>"
							sOutputText = sOutputText & "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							sOutputText = sOutputText & "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(sLoadError)) & "</FONT></TD>"
							sOutputText = sOutputText &  "</TR>"
							lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
						End If
					Next
					sOutputText = ""
					sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de número de empleado:&nbsp;" & CleanStringForHTML(CStr(lRFCError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de estatus del empleado:&nbsp;" & CleanStringForHTML(CStr(lEmployeeStatusError)) & "</FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha inicial:&nbsp;" & CleanStringForHTML(CStr(lStartDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de fecha final:&nbsp;" & CleanStringForHTML(CStr(lDateError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de clave de concepto:&nbsp;" & CleanStringForHTML(CStr(lConceptError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores de registro duplicado:&nbsp;" & CleanStringForHTML(CStr(lCreditError)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Errores al agregar el crédito:&nbsp;" & CleanStringForHTML(CStr(lAddError)) & "</B></FONT></TD></TR>"
					'sOutputText = sOutputText & "<TR></TR><TR></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros :&nbsp;" & CleanStringForHTML(CStr(lTotal)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros correctos:&nbsp;" & CleanStringForHTML(CStr(lTotalSucess)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Altas:&nbsp;" & CleanStringForHTML(CStr(lA)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Cambios:&nbsp;" & CleanStringForHTML(CStr(lC)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>--> Bajas:&nbsp;" & CleanStringForHTML(CStr(lB)) & "</B></FONT></TD></TR>"
					sOutputText = sOutputText & "<TR><TD COLSPAN=""4"" ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2""><B>Total de registros rechazados:&nbsp;" & CleanStringForHTML(CStr(lTotalError)) & "</B></FONT></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileReportName, sOutputText, sErrorDescription)
				End If
				lErrorNumber = AppendTextToFile(sFileReportName, "<!-- FileName:" & sFileName & "-->", sErrorDescription)
				sTarget = Replace(sFileReportName, SYSTEM_PHYSICAL_PATH, "", 1, 1, vbBinaryCompare)
				sTarget = "<A HREF="" " & sTarget & """ target=""_blank"">Ver informe completo de carga</A>"
				Call DisplayErrorMessage("La información ha sido procesada", sTarget)
		End Select
	Else
		Call DisplayErrorMessage("Error al registrar la información", "<TR> <TD> YA EXISTEN REGISTROS DEL ARCHIVO A CARGAR </TD> </TR>")
	End If
End Function

Public Function VerifyExistenceOfFileRegisters(sFileName, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To verify if exist records from file uploaded
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfFileRegisters"
	Dim lErrorNumber
	Dim oRecordset

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Credits Where (UploadedFileName = '" & sFileName & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If (lErrorNumber = 0) Then
		If (NOT oRecordset.EOF) Then
			aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditID").Value)
			VerifyExistenceOfFileRegisters = true
		Else
			aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
			VerifyExistenceOfFileRegisters = false
		End If
	Else
		sErrorDescription = "Error al verificar la existencia de registros del archivo indicado."
		aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
		VerifyExistenceOfFileRegisters = false
	End If
	Err.Clear
End Function

Public Function VerifyIfUploadMonthDateIsCorrect(lDate, sErrorDescription)
'************************************************************
'Purpose: To verify if exist records from file uploaded
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyIfUploadMonthDateIsCorrect"
	Dim lMonth

	If (CLng(lDate) = 30000000) Or (CLng(lDate) = 0) Then
		VerifyIfUploadMonthDateIsCorrect = True
	Else
		lMonth = Mid(CStr(lDate), Len("00000"), Len("00"))
		If (CInt(lMonth)>0) And (CInt(lMonth)<13) Then
			VerifyIfUploadMonthDateIsCorrect = True
		Else
			sErrorDescription = "El mes de la fecha " & DisplayNumericDateFromSerialNumber(lDate) & " indicada es incorrecto."
			VerifyIfUploadMonthDateIsCorrect = False
		End If
	End If
End Function

Function UploadHistNomsarFile(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To insert each entry in the given file into the
'         DM_Hist_Nomsar table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadHistNomsarFile"
	Dim sFileContents
	Dim asFileContents
	Dim iIndex
	Dim jIndex
	Dim sErrorQueries
	Dim lErrorNumber
	Dim asFileRow
	Dim sQuery
	Dim sDate
	Dim lCurrentDate
	Dim lPeriodID
	Dim oRecordset
	Dim lNextID
	Dim lMinDate
	Dim lMaxDate
	
	lErrorNumber = GetNewIDFromTable(oADODBConnection, "dm_sar_periods", "PeriodID", "", 1, lNextID, sErrorDescription)	
	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	sQuery = "Select Count(*) As Total From Dm_Sar_Periods Where (IsOpen=1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If CLng(oRecordset.Fields("Total").Value) <> 0 Then
		sFileContents = GetFileContents(sFileName, sErrorDescription)
		If Len(sFileContents) > 0 Then
			asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
			sErrorQueries = ""
			asFileRow = Split(asFileContents(0), vbTab, -1, vbBinaryCompare)
			For iIndex = 0 To UBound(asFileRow)
				'If StrComp(oRequest("Column" & (iIndex + 1)),"PeriodID",vbBinaryCompare) = 0 Then lPeriodID = Mid(asFileRow(iIndex), 2) & "0" & Mid(asFileRow(iIndex), 1, 1)
                If StrComp(oRequest("Column" & (iIndex + 1)),"PeriodID",vbBinaryCompare) = 0 Then lPeriodID = Mid(asFileRow(iIndex), 3) & "0" & Mid(asFileRow(iIndex), 2, 1) 
			Next
			sQuery = "Select IsOpen From Dm_Sar_Periods Where (PeriodName=" & lPeriodID & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If Not oRecordset.EOF Then
				If CInt(oRecordset.Fields("IsOpen").Value) = 1 Then
					sQuery = "Select Count(*) Tot From Dm_Hist_Nomsar Where PeriodID = " & CLng(lPeriodID)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If oRecordset.Fields("Tot").Value = 0 Then
						For iIndex = 0 To UBound(asFileContents)-1
							sQuery = "Insert Into DM_Hist_Nomsar ("
							asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
							For jIndex = 0 To UBound(asFileRow)
								If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
									'If StrComp(oRequest("Column" & (jIndex + 1)),"PeriodID",vbBinaryCompare) = 0 Then lPeriodID = Mid(asFileRow(jIndex), 2) & "0" & Mid(asFileRow(jIndex), 1, 1)
									If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
										(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
										(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
										If jIndex < UBound(asFileRow) Then
											sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ","
										Else
											sQuery = sQuery & sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ",UserID,LastUpdateDate) Values ("
										End If
									Else
										If jIndex < UBound(asFileRow) Then
											sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ","
										Else
											sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ",UserID,LastUpdateDate) Values ("
										End If
									End If
								End If
							Next
							For jIndex = 0 To UBound(asFileRow)
								If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
									If StrComp(oRequest("Column" & (jIndex + 1)),"PeriodID",vbBinaryCompare) = 0 Then
										sQuery = sQuery & lPeriodID & ","
									ElseIf (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
										(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
										(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
											If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Then sDate = asFileRow(jIndex)
											If (InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),1,2) & Mid(asFileRow(jIndex),4,2)
											If (InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),4,2) & Mid(asFileRow(jIndex),1,2)
											If Len(sDate) = 0 Then sDate = "0"
											sQuery = sQuery & sDate & ","
									ElseIf StrComp(oRequest("Column" & (jIndex + 1)),"CLC",vbBinaryCompare) = 0 Then
										sQuery = sQuery & "'x" & asFileRow(jIndex) & "',"
									Else
										If isNumeric(asFileRow(jIndex)) Then
											If jIndex < UBound(asFileRow) Then
												sQuery = sQuery & asFileRow(jIndex) & ","
											Else
												sQuery = sQuery & asFileRow(jIndex) & "," & aLoginComponent(N_USER_ID_LOGIN) & "," & lCurrentDate & ")"
											End If
										Else
											If jIndex < UBound(asFileRow) Then
												sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "',"
											Else
												sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "'," & aLoginComponent(N_USER_ID_LOGIN) & "," & lCurrentDate & ")"
											End If
										End If
									End If
								End If
							Next
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						Next
					Else
						lErrorNumber = -1
						sErrorQueries = "Ya existen registros previos correspondientes al periodo " & lPeriodID
					End If
					sErrorDescription = sErrorQueries
				Else
					sErrorDescription="El periodo correspondiente al reporte que intenta cargar está cerrado, el proceso no puede continuar"
					lErrorNumber = -1
				End If
			Else
				sErrorDescription="El periodo correspondiente al reporte que intenta cargar no ha existe"
				lErrorNumber = -1
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se han encontrado periodos abiertos, la carga no puede continuar"
	End If
	sErrorDescription = sErrorQueries
	UploadHistNomsarFile = lErrorNumber
End Function

Function UploadBanamexCensus(oADODBConnection, sFileName, sErrorDescription)
'************************************************************
'Purpose: To load the Banamex Census Info.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadBanamexCensus"
	Dim sFileContents
	Dim asFileContents
	Dim iIndex
	Dim jIndex
	Dim iEmployeeIndex
	Dim sErrorQueries
	Dim lErrorNumber
	Dim asFileRow
	Dim sQuery
	Dim sDate
	Dim sQueryCompany
	Dim lCurrentDate
	Dim oRecordset
	Dim iCompanyID
	Dim sTable
	
	If StrComp(oRequest("Load").Item,"SarCensus",VbBinaryCompare) = 0 Then
		sTable = "DM_PADRON_BANAMEX_NUEVO"
	Else
		sTable = "DM_PADRON_BANAMEX"
	End If
	
	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	sFileContents = GetFileContents(sFileName, sErrorDescription)
	If Len(sFileContents) > 0 Then
		asFileContents = Split(sFileContents, vbNewLine, -1, vbBinaryCompare)
		sErrorQueries = ""
		sQuery = "Truncate Table " & sTable
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		For iIndex = 0 To UBound(asFileContents)
			sQuery = "Insert Into " & sTable & " ("
			asFileRow = Split(asFileContents(iIndex), vbTab, -1, vbBinaryCompare)
			For jIndex = 0 To UBound(asFileRow)
				If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
					If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
						If jIndex < UBound(asFileRow) Then
							sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ","
						Else
							sQuery = sQuery & sQuery = sQuery & Mid(oRequest("Column" & (jIndex + 1)),1,Len(oRequest("Column" & (jIndex + 1)))-8) & ",UserID,LastUpdateDate) Values ("
						End If
					Else
						If jIndex < UBound(asFileRow) Then
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ","
						Else
							sQuery = sQuery & oRequest("Column" & (jIndex + 1)) & ",UserID,LastUpdateDate) Values ("
						End If
					End If
				End If
			Next
			For jIndex = 0 To UBound(asFileRow)
				If (StrComp(oRequest("Column" & (jIndex + 1)),"NA",vbBinaryCompare) <> 0) Then
					If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Or _
						(InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "YYYYMMDD", vbBinaryCompare) > 0) Then sDate = asFileRow(jIndex)
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "MMDDYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),1,2) & Mid(asFileRow(jIndex),4,2)
							If (InStr(1, oRequest("Column" & (jIndex + 1)), "DDMMYYYY", vbBinaryCompare) > 0) Then sDate = Mid(asFileRow(jIndex),7,4) & Mid(asFileRow(jIndex),4,2) & Mid(asFileRow(jIndex),1,2)
							If Len(sDate) = 0 Then sDate = "0"
							sQuery = sQuery & sDate & ","
					Else
					    If (oRequest("Column" & (jIndex + 1)).item = "mot_baja") AND (asFileRow(jIndex)="") Then asFileRow(jIndex)="0"
						If isNumeric(asFileRow(jIndex)) Then
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & asFileRow(jIndex) & ","
							Else
								sQuery = sQuery & asFileRow(jIndex) & "," & aLoginComponent(N_USER_ID_LOGIN) & "," & lCurrentDate & ")"
							End If
						Else
							If jIndex < UBound(asFileRow) Then
								sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "',"
							Else
								sQuery = sQuery & "'" & Trim(asFileRow(jIndex)) & "'," & aLoginComponent(N_USER_ID_LOGIN) & "," & lCurrentDate & ")"
							End If
						End If
					End If
				End If
			Next
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		Next

        sQuery="Update " & sTable & " Set u_version='2' Where (u_version='S') "
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

		If iConnectionType <> ORACLE Then
		   sQuery = "Update " & sTable & " Set u_version = (Select Cast(CompanyID As Varchar) From Employees Where (EmployeeID = " & sTable & ".EmployeeID) And (CompanyId <> -1) AND " &  sTable & ".u_version<>'2' )"
        Else
           sQuery = "Update " & sTable & " Set u_version = (Select To_Char(CompanyID) From Employees Where (EmployeeID = " & sTable & ".EmployeeID) And (CompanyId <> -1) AND " &  sTable & ".u_version<>'2' )" 
        End If

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		sErrorDescription = sErrorQueries

        sQuery="Update " & sTable & " Set u_version='0' where (u_version Is Null)"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If	
End Function

Function UpdateHistoryForward(oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: Updates job and employee history and current
'		  employee information if job is taken
'         DM_Hist_Nomsar table.
'Inputs:  oADODBConnection, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UploadHistNomsarFile"
	Dim sQuery
	Dim oRecordset
	Dim lErrorNumber
	
	sQuery = "Select StatusID From Jobs Where JobID = 35156"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If CInt(oRecordset.Fields("StatusID").Value) = 2 Then
		sQuery = "Update Employees Set PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ",ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ",GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ",IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ",LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")"
		sErrorDescription = "No se pudo actualizar la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sQuery = "Update EmployeesHistoryList Set PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ",ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ",GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ",IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ",LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ",PositionID=" & aJobComponent(N_POSITION_ID_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")"
		sErrorDescription = "No se pudo actualizar el historial de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sQuery = "Update JobsHistoryList Set PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ",ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ",GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ",IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ",LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "UploadInfoLibrary.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			If lErrorNumber <> 0 Then sErrorDescription = "No se pudo actuallizar el historial de la plaza"
		End If
	End If
	UpdateHistoryForward = lErrorNumber
End Function
%>