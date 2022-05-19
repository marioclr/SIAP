<%

Function GetEmployeesURLValues(oRequest, iSelectedTab, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: iSelectedTab, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesURLValues"
	Dim oItem
	Dim aItem

	iSelectedTab = 1
	If Not IsEmpty(oRequest("Tab").Item) Then
		iSelectedTab = CInt(oRequest("Tab").Item)
	End If
	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0) Or (Len(oRequest("AuthorizationFile").Item) > 0)

	sCondition = ""
	If Len(oRequest("EmployeeNumber").Item) > 0 Then
		sCondition = sCondition & " And (EmployeeNumber Like ('" & S_WILD_CHAR & Replace(oRequest("EmployeeNumber").Item, "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("EmployeeName").Item) > 0 Then
		sCondition = sCondition & " And (EmployeeName Like '" & S_WILD_CHAR & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR), "Á", S_WILD_CHAR), "á", S_WILD_CHAR), "É", S_WILD_CHAR), "é", S_WILD_CHAR), "Í", S_WILD_CHAR), "í", S_WILD_CHAR), "Ó", S_WILD_CHAR), "ó", S_WILD_CHAR), "Ú", S_WILD_CHAR), "ú", S_WILD_CHAR), "Ñ", S_WILD_CHAR), "ñ", S_WILD_CHAR) & S_WILD_CHAR & "')"
	End If
	If Len(oRequest("EmployeeLastName").Item) > 0 Then
		sCondition = sCondition & " And (EmployeeLastName Like '" & S_WILD_CHAR & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(oRequest("EmployeeLastName").Item, "'", S_WILD_CHAR), "Á", S_WILD_CHAR), "á", S_WILD_CHAR), "É", S_WILD_CHAR), "é", S_WILD_CHAR), "Í", S_WILD_CHAR), "í", S_WILD_CHAR), "Ó", S_WILD_CHAR), "ó", S_WILD_CHAR), "Ú", S_WILD_CHAR), "ú", S_WILD_CHAR), "Ñ", S_WILD_CHAR), "ñ", S_WILD_CHAR) & S_WILD_CHAR & "')"
	End If
	If Len(oRequest("EmployeeLastName2").Item) > 0 Then
		sCondition = sCondition & " And (EmployeeLastName2 Like '" & S_WILD_CHAR & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(oRequest("EmployeeLastName2").Item, "'", S_WILD_CHAR), "Á", S_WILD_CHAR), "á", S_WILD_CHAR), "É", S_WILD_CHAR), "é", S_WILD_CHAR), "Í", S_WILD_CHAR), "í", S_WILD_CHAR), "Ó", S_WILD_CHAR), "ó", S_WILD_CHAR), "Ú", S_WILD_CHAR), "ú", S_WILD_CHAR), "Ñ", S_WILD_CHAR), "ñ", S_WILD_CHAR) & S_WILD_CHAR & "')"
	End If
	If Len(oRequest("RFC").Item) > 0 Then
		sCondition = sCondition & " And (RFC In ('" & Replace(oRequest("RFC").Item, ", ", ",") & "'))"
	End If
	If (InStr(1, oRequest, "StartBirth", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndBirth", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartBirth", "EndBirth", "BirthDate", False, sCondition)
	If Len(oRequest("CompanyID").Item) > 0 Then
		sCondition = sCondition & " And (Areas.CompanyID In (" & Replace(oRequest("CompanyID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("EmployeeTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.EmployeeTypeID In (" & Replace(oRequest("EmployeeTypeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("PositionTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.PositionTypeID In (" & Replace(oRequest("PositionTypeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("ClassificationID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.ClassificationID=" & oRequest("ClassificationID").Item & ")"
	End If
	If Len(oRequest("GroupGradeLevelID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.GroupGradeLevelID In (" & Replace(oRequest("GroupGradeLevelID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("IntegrationID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.IntegrationID=" & oRequest("IntegrationID").Item & ")"
	End If
	If Len(oRequest("LevelID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.LevelID In (" & Replace(oRequest("LevelID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("JourneyID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.JourneyID In (" & Replace(oRequest("JourneyID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("ShiftID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.ShiftID In (" & Replace(oRequest("ShiftID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("WorkingHours").Item) > 0 Then
		sCondition = sCondition & " And (Employees.WorkingHours=" & oRequest("WorkingHours").Item & ")"
	End If
	If Len(oRequest("StatusID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.StatusID In (" & Replace(oRequest("StatusID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("JobID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.JobID =" & oRequest("JobID").Item & ")"
	End If
	If Len(oRequest("JobTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Jobs.JobTypeID =" & oRequest("JobTypeID").Item & ")"
	End If
	If Len(oRequest("DoSearch").Item) > 0 Then Response.Cookies("SIAP_SearchPath").Item = oRequest

	GetEmployeesURLValues = Err.number
	Err.Clear
End Function

Function DoEmployeesAction(oRequest, oADODBConnection, sAction, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoEmployeesAction"
	Dim oRecordset
	Dim sNames
	Dim oItem
	Dim lErrorNumber
	Dim lEmployeeID
	Dim iIndex
	Dim sErrorDescription1
	Dim sErrorDescription2
	Dim sAbsenceIDs
	Dim iAbsenceID
	Dim iActiveOriginal
	Dim sAbsenceType
	Dim iReasonID
	Dim lToday

	If Len(oRequest("RemoveFile").Item) > 0 Then
		If FileExists(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription) Then
			lErrorNumber = DeleteFile(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription)
		End If
	ElseIf Len(oRequest("Add").Item) > 0 Then
		Select Case sAction
			Case "PayrollRevision"
				lErrorNumber = AddEmployeeRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
			Case "Revisiones"
				Dim lCont
				Dim lPayrollYear
				Dim lAmount
				Dim lStartDateForRevision
				For Each oItem In oRequest("PayrollRevision")
					 aPayrollRevisionComponent(N_START_DATE_REVISION) = CLng(oItem)
					 lPayrollYear = CInt(Left(aPayrollRevisionComponent(N_START_DATE_REVISION), Len("0000")))
					 Select Case aPayrollRevisionComponent(N_CONCEPT_ID_REVISION)
						Case 5, 19, 22, 24, 26, 32, 45, 46, 50, 69, 63, 67, 70, 72, 73, 76, 77, 93, 94, 104, 146
							' Para estos conceptos se revisa si estan activos con su fecha de inicio y fin
							If VerifyConceptInPayrrollForEmployee(oRequest, oADODBConnection, aPayrollRevisionComponent, lPayrollYear, lAmount, sErrorDescription) Then
								If Not VerifyConceptIsActiveInPeriod(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription) Then
									' Se agrega un registro negativo para descuento de concepto que no debio de haberse pagado
									aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = lAmount * (-1)
									lErrorNumber = AddEmployeeAdjustmentForRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription1 = sErrorDescription1 & " " & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
									End If
								End If
							Else
								If VerifyConceptIsActiveInPeriod(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription) Then
									' Se agrega un registro positivo para pago de concepto que debio de haberse pagado
									aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = lAmount
									lErrorNumber = AddEmployeeAdjustmentForRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription2 = sErrorDescription2 & " " & CStr(iAbsenceID) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
									End If
								End If
							End If
						Case 9, 14
							' Para estos conceptos se revisa si estan activos recorriendo todos los días del periodo analizado
							lStartPayrollDate = GetPayrollStartDate(aPayrollRevisionComponent(N_START_DATE_REVISION))
							For lCont = lStartPayrollDate To aPayrollRevisionComponent(N_START_DATE_REVISION)
								aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) = lCont
								' Para estos conceptos se revisa si estan activos con su fecha de inicio y fin
								If VerifyConceptInPayrrollForEmployee(oRequest, oADODBConnection, aPayrollRevisionComponent, lPayrollYear, lAmount, sErrorDescription) Then
									If Not VerifyConceptIsActiveInPeriod(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription) Then
										' Se agrega un registro negativo para descuento de concepto que no debio de haberse pagado
										aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = lAmount * (-1)
										lErrorNumber = AddEmployeeAdjustmentForRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
										If lErrorNumber = 0 Then
											sErrorDescription1 = sErrorDescription1 & " " & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
										End If
									End If
								Else
									If VerifyConceptIsActiveInPeriod(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription) Then
										' Se agrega un registro positivo para pago de concepto que debio de haberse pagado
										aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = lAmount
										lErrorNumber = AddEmployeeAdjustmentForRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
										If lErrorNumber = 0 Then
											sErrorDescription2 = sErrorDescription2 & " " & CStr(iAbsenceID) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
										End If
									End If
								End If
							Next
					 End Select
				Next
				If Len(sErrorDescription1) > 0 Then
					lErrorNumber = -1
					sErrorDescription1 = Left(sErrorDescription1, (Len(sErrorDescription1) - Len(","))) & "."
					sErrorDescription = "Se descontara lo siguiente por revisión: " & sErrorDescription1
				End If
				If Len(sErrorDescription2) > 0 Then
					sErrorDescription2 = Left(sErrorDescription2, (Len(sErrorDescription2) - Len(","))) & "."
					sErrorDescription = "Se recuperara lo siguiente por revisión: " & sErrorDescription2 & "</BR>" & sErrorDescription
				End If
			Case "Absences"
				If Len(oRequest("CancelAbsence").Item) > 0 Then
					lErrorNumber = AddJustification(oRequest, oADODBConnection, aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), iActiveOriginal, aAbsenceComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						sErrorDescription2 = sErrorDescription2 & " " & CStr(iAbsenceID) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
					End If
				ElseIf Len(oRequest("Justification").Item) > 0 Then
					If VerifyExistenceOfAbsencesForJustification(oADODBConnection, aAbsenceComponent, aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), false, iAbsenceID, iActiveOriginal, sErrorDescription) Then
						If iActiveOriginal = 0 Then
							aAbsenceComponent(N_ACTIVE_ABSENCE) = -1
						Else
							aAbsenceComponent(N_ACTIVE_ABSENCE) = -2
						End If
						lErrorNumber = AddJustification(oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							sErrorDescription2 = sErrorDescription2 & " " & CStr(iAbsenceID) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
						End If
					Else
						lErrorNumber = -1
					End If
				Else
					Call VerifyAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceType, sErrorDescription)
					If Len(oRequest("OcurredDates").Item) > 0 Then
						'If VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
						If VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) And ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>21) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>22) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)<>23)) Then
						' Es Incidencia por periodo, tengo fecha inicio y fecha fin.
						' Validar si inicio de incidencia no cae en Licencia/Suspensión
							For iIndex = 1 To 2
								If iIndex = 1 Then aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(oRequest("OcurredDates")(1))
								If iIndex = 2 Then
									Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
										Case 41, 42, 43, 44, 45, 46, 47, 48, 49, 57, 58
											aAbsenceComponent(N_END_DATE_ABSENCE) = CLng(oRequest("OcurredDates")(2))
											If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = aAbsenceComponent(N_END_DATE_ABSENCE) Then
												aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
											End If
										Case 50, 51, 54, 55, 56
											aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000
										Case Else
											aAbsenceComponent(N_END_DATE_ABSENCE) = CLng(oRequest("OcurredDates")(2))
									End Select
								End If
							Next
							aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
							aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aAbsenceComponent(N_END_DATE_ABSENCE)
							If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								Select Case sAbsenceType
									Case "Suspensions"
										lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
										If lErrorNumber = 0 Then
											aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 159
											Call GetReasonIDfromAbsence(iReasonID, sErrorDescription)
											If lErrorNumber = 0 Then
												aEmployeeComponent(N_REASON_ID_EMPLOYEE) = iReasonID
												lErrorNumber = ModifyEmployeeForSuspension(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
												If lErrorNumber <> 0 Then
													sErrorDescription = "Error al actualizar los datos del empleado."
												End If
											Else
												sErrorDescription = "No se pudo obtener el motivo para registrar la suspensión."
											End If
										Else
											sErrorDescription = "No se pudo obtener la información del empleado para registrar la suspensión."
										End If
									Case Else
										lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
								End Select
							Else
								lErrorNumber = -1
							End If
						Else
							Select Case sAbsenceType
								Case "Justification"
									lErrorNumber = GetAbsenceAppliesToID(oRequest, oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
									For Each oItem In oRequest("OcurredDates")
										aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(oItem)
										aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
											If (VerifyExistenceOfAbsencesForJustification(oADODBConnection, aAbsenceComponent, aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), false, iAbsenceID, iActiveOriginal, sErrorDescription)) Then
												If iActiveOriginal = 0 Then
													aAbsenceComponent(N_ACTIVE_ABSENCE) = -1
												Else
													aAbsenceComponent(N_ACTIVE_ABSENCE) = -2
												End If
												lErrorNumber = AddJustification(oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
												If lErrorNumber = 0 Then
													sErrorDescription2 = sErrorDescription2 & " " & CStr(iAbsenceID) & " del día " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
												End If
											Else
												sErrorDescription1 = sErrorDescription1 & " " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ","
											End If
										Else
											lErrorNumber = -1
											'sErrorDescription = "El empleado no esta activo dentro del periodo de la incidencia capturada."
										End If
									Next
									If Len(sErrorDescription1) > 0 Then
										lErrorNumber = -1
										sErrorDescription1 = Left(sErrorDescription1, (Len(sErrorDescription1) - Len(","))) & "."
										sErrorDescription = "No existe incidencia a justificar para el " & sErrorDescription1
									End If
									If Len(sErrorDescription2) > 0 Then
										sErrorDescription2 = Left(sErrorDescription2, (Len(sErrorDescription2) - Len(","))) & "."
										sErrorDescription = "Se justificaron las incidencias " & sErrorDescription2 & "</BR>" & sErrorDescription
									End If
								Case "Suspension"
								Case Else
									For Each oItem In oRequest("OcurredDates")
										aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(oItem)
										aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
										If VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
											lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												sErrorDescription1 = sErrorDescription1 & " " & sErrorDescription & ";"
											End If
										Else
											lErrorNumber = -1
											sErrorDescription2 = sErrorDescription2 & " " & DisplayDateFromSerialNumber(CLng(oItem), -1, -1, -1) & ";"
										End If
									Next
									sErrorDescription = ""
									If Len(sErrorDescription1) > 0 Then
										lErrorNumber = -1
										sErrorDescription1 = Left(sErrorDescription1, (Len(sErrorDescription1) - Len(","))) & "."
										sErrorDescription = "Errores al agregar:" & sErrorDescription1
									End If
									If Len(sErrorDescription2) > 0 Then
										lErrorNumber = -1
										sErrorDescription2 = Left(sErrorDescription2, (Len(sErrorDescription2) - Len(","))) & "."
										sErrorDescription2 = "Empleado inactivo en el periodo de la(s) incidencia(s)" & sErrorDescription2 & "</BR>"
										sErrorDescription  = sErrorDescription & "</BR>" & sErrorDescription2
									End If
							End Select
						End If
					Else
						lErrorNumber = AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					End If
				End If
				'Redim aAbsenceComponent(N_ABSENCE_COMPONENT_SIZE)
				'aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = aEmployeeComponent(N_ID_EMPLOYEE)
				'aAbsenceComponent(N_HOURS_ABSENCE) = 0
			Case "ChildrenSchoolarships"
				lErrorNumber = SaveEmployeeChildren(aEmployeeComponent, "ChildrenSchoolarships", sErrorDescription)
			Case "DocumentsForLicenses"
				lErrorNumber = AddDocumentsForLicenses(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeesBeneficiaries"
				lErrorNumber = AddEmployeeBeneficiary(oRequest, oADODBConnection, "", aEmployeeComponent, sErrorDescription)
			Case "EmployeesChildren"
				lErrorNumber = SaveEmployeeChildren(aEmployeeComponent, "EmployeesChildren", sErrorDescription)
			Case "EmployeeConcepts"
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
				lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeeHistoryList"
				lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
			Case "EmployeePayroll"
				lErrorNumber = ModifyEmployeePayroll(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "Employees"
				lErrorNumber = AddEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lErrorNumber <> 0 Then aEmployeeComponent(N_ID_EMPLOYEE) = -1
		End Select
	ElseIf Len(oRequest("AuthorizationFile").Item) > 0 Then
		Select Case sAction
			Case "ApplyAbsences"
				If (Len(oRequest("OnlyAttendanceControl").Item) > 0) Or (Len(oRequest("OnlySuspension").Item) > 0) Then
					If Len(oRequest("OnlyAttendanceControl").Item) > 0 Then lErrorNumber = ApplyAttendanceControlInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					If Len(oRequest("OnlySuspension").Item) > 0 Then lErrorNumber = ApplySuspensionsInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
				Else
					lErrorNumber = ApplyAttendanceControlInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					lErrorNumber = ApplySuspensionsInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = ApplyAbsencesInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					End If
				End If
		End Select
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		Select Case sAction
			Case "Absences"
				If Len(oRequest("Justification").Item) > 0 Then
					iActiveOriginal = CInt(oRequest("Active").Item)
					If iActiveOriginal = 0 Then
						aAbsenceComponent(N_ACTIVE_ABSENCE) = -1
					Else
						aAbsenceComponent(N_ACTIVE_ABSENCE) = -2
					End If
					lErrorNumber = ModifyAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
				ElseIf Len(oRequest("CancelJustification").Item) > 0 Then
					lErrorNumber = CancelJustification(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
				ElseIf Len(oRequest("CancelAbsence").Item) > 0 Then
					lErrorNumber = CancelAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
					'Redim aAbsenceComponent(N_ABSENCE_COMPONENT_SIZE)
					aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = aEmployeeComponent(N_ID_EMPLOYEE)
				End If
			Case "ChildrenSchoolarships"
				lErrorNumber = ModifyEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeesNew"
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
					lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
					End If
				End If
				Response.Redirect "Employees.asp" & "?Action=EmployeesNew&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & "&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&AssignJob=1"
			Case "EmployeesBeneficiaries"
				lErrorNumber = ModifyEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeesChildren"
				lErrorNumber = ModifyEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeeConcepts"
				lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeeHistoryList"
				lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 2, aEmployeeComponent, sErrorDescription)
			Case "EmployeeJob"
				If Len(oRequest("JobID").Item) > 0 Then
					lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						aEmployeeComponent(N_JOB_ID_EMPLOYEE) = CLng(oRequest("JobID").Item)
						lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
				End If
			Case "CheckJobNumber"
				If Len(oRequest("JobID").Item) > 0 Then
					aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRequest("EmployeeID").Item)
					aEmployeeComponent(N_JOB_ID_EMPLOYEE) = CLng(oRequest("JobID").Item)
					aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = -3
					aEmployeeComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE) = False
					aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = False
					lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				End If
			Case "Employees"
				lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						aJobComponent(N_SERVICE_ID_JOB) = aEmployeeComponent(N_SERVICE_ID_EMPLOYEE)
						aJobComponent(N_LEVEL_ID_JOB) = aEmployeeComponent(N_LEVEL_ID_EMPLOYEE)
						lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					End If
				End If
			Case "Jobs"
				lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "SaveEmployeeError"
				lErrorNumber = AddEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("RemoveValidate").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			If lErrorNumber = 0 Then
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = DropEmployee(oRequest, oADODBConnection, 2, aEmployeeComponent, sErrorDescription)
				End If
			End If
		End If
		Response.Redirect "Employees.asp?Pending=1"
	ElseIf Len(oRequest("RemoveMotion").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			If lErrorNumber = 0 Then
				lErrorNumber = RemoveEmployeeForValidation(oADODBConnection, aEmployeeComponent, sErrorDescription)
			End If
		End If
		Response.Write "Movimiento" & CLng(oRequest("ReasonID").Item)
		Response.Redirect "UploadInfo.asp?Action=" & sAction & "&ReasonID=" & CLng(oRequest("ReasonID").Item)
	ElseIf Len(oRequest("RemoveAuthorization").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			If lErrorNumber = 0 Then
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = DropEmployee(oRequest, oADODBConnection, 3, aEmployeeComponent, sErrorDescription)
				End If
			End If
		End If
		Response.Redirect "Employees.asp?Pending=1"
	ElseIf Len(oRequest("RemoveApply").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			If lErrorNumber = 0 Then
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = DropEmployee(oRequest, oADODBConnection, -1, aEmployeeComponent, sErrorDescription)
				End If
			End If
		End If
		If lErrorNumber = 0 Then
			Response.Redirect "UploadInfo.asp?Action=EmployeesDrop&Success=1&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE)
		Else
			Response.Redirect "UploadInfo.asp?Action=EmployeesDrop&Success=0&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE)
		End If
	ElseIf Len(oRequest("LicenseValidate").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = AddLicenseEmployee(oRequest, oADODBConnection, 2, aEmployeeComponent, sErrorDescription)
				End If
			End If
		End If
	ElseIf Len(oRequest("LicenseAuthorization").Item) > 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
			lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				lErrorNumber = AddLicenseEmployee(oRequest, oADODBConnection, 3, aEmployeeComponent, sErrorDescription)
			End If
		End If
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		Select Case sAction
			Case "Absences"
				lErrorNumber = RemoveAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
				'Redim aAbsenceComponent(N_ABSENCE_COMPONENT_SIZE)
				aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = aEmployeeComponent(N_ID_EMPLOYEE)
			Case "ChildrenSchoolarships"
				lErrorNumber = RemoveEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1
			Case "EmployeesBeneficiaries"
				lErrorNumber = RemoveEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1
			Case "EmployeesChildren"
				lErrorNumber = RemoveEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1
			Case "EmployeeConcepts"
				lErrorNumber = RemoveEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "EmployeeHistoryList"
				lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 0, aEmployeeComponent, sErrorDescription)
			Case "Employees"
				If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
					lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = RemoveEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then aEmployeeComponent(N_ID_EMPLOYEE) = -1
						Response.Redirect "UploadInfo.asp?Action=EmployeesNew"
					End If
				End If
			Case "PayrollRevision"
				lErrorNumber = RemoveEmployeeRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("ResumptionOfWorkValidate").Item) > 0 Then
		aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 2
		lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = AddResumptionOfWork(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End If
	ElseIf Len(oRequest("ResumptionOfWorkAuthorization").Item) > 0 Then
		aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 3
		lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = AddResumptionOfWork(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End If
	ElseIf Len(oRequest("ResumptionOfWorkApply").Item) > 0 Then
		aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 0
		lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = AddResumptionOfWork(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End If
	ElseIf Len(oRequest("Register").Item) > 0 Then
		If Len(oRequest("SaveAssignJob").Item) > 0 Then
			lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
			lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = SetRegisterForEmployee(oRequest, oADODBConnection, 2, aEmployeeComponent, sErrorDescription)
			End If
		End If
		Response.Redirect "Employees.asp?Pending=1"
	ElseIf Len(oRequest("Validate").Item) > 0 Then
		If Len(oRequest("SaveAssignJob").Item) > 0 Then
			lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
				lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
				End If
			End If
		Else 
			If lErrorNumber = 0 Then
				aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
				lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
				End If
			End If
		End If
		lErrorNumber = SetRegisterForEmployee(oRequest, oADODBConnection, 3, aEmployeeComponent, sErrorDescription)
		Response.Redirect "Employees.asp?Pending=1"
	ElseIf Len(oRequest("Authorization").Item) > 0 Then
		If Len(oRequest("SaveAssignJob").Item) > 0 Then
			lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
				lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
				End If
			End If
		Else
			If Len(oRequest("ReasonType").Item) > 0 Then
				Select Case CInt(oRequest("ReasonType").Item)
					Case 3
						If lErrorNumber = 0 Then
							lErrorNumber = DropEmployee(oRequest, oADODBConnection, -1, aEmployeeComponent, sErrorDescription)
						End If
					Case 4, 5
						If lErrorNumber = 0 Then
							lErrorNumber = AddLicenseEmployee(oRequest, oADODBConnection, -2, aEmployeeComponent, sErrorDescription)
						End If
					Case 6
						If lErrorNumber = 0 Then
							lErrorNumber = AddLicenseEmployee(oRequest, oADODBConnection, -3, aEmployeeComponent, sErrorDescription)
						End If
					Case Else
						If lErrorNumber = 0 Then
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE)
							lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								lErrorNumber = ModifyEmployeeJob(oRequest, oADODBConnection, aEmployeeComponent, aJobComponent, sErrorDescription)
							End If
							lErrorNumber = SetRegisterForEmployee(oRequest, oADODBConnection, 0, aEmployeeComponent, sErrorDescription)
						End If
				End Select
				Response.Redirect "Employees.asp?Pending=1"
			End If
		End If
	ElseIf Len(oRequest("SetActive").Item) > 0 Then
		Select Case sAction
			Case "EmployeeConcepts"
				lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "Employees"
				lErrorNumber = SetActiveForEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Case "Absences"
				lErrorNumber = SetActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		End Select
	'ElseIf Len(oRequest("AssignJob").Item) > 0 Then
	'	Select Case sAction
	'		Case "EmployeesNew"
	'			Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), ",AssignJob,", ",1,4,5,", lReasonID, aEmployeeComponent, sErrorDescription)
	'	End Select
	End If

	Set oRecordset = Nothing
	DoEmployeesAction = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesSearchForm(oRequest, oADODBConnection, sAction, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesSearchForm"
		If StrComp(sAction,"Catalogs.asp",vbBinaryCompare) = 0 Then
			Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & sAction & """ METHOD=""GET"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""ActionHdn"" VALUE=""2"" />"
		Else
			Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & sAction & """ METHOD=""GET"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ChildrenSchoolarships"" />"
		End If
		If bFull Then Response.Write "<B>BÚSQUEDA DE EMPLEADOS</B><BR /><BR />"
		Response.Write "<TABLE"
			If Not bFull Then Response.Write " WIDTH=""400"""
		Response.Write " BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""7"" MAXLENGTH=""7"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""25"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" SIZE=""25"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" SIZE=""25"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If bFull Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("StartBirthYear").Item), CInt(oRequest("StartBirthMonth").Item), CInt(oRequest("StartBirthDay").Item), "StartBirthYear", "StartBirthMonth", "StartBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("EndBirthYear").Item), CInt(oRequest("EndBirthMonth").Item), CInt(oRequest("EndBirthDay").Item), "EndBirthYear", "EndBirthMonth", "EndBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" SIZE=""13"" MAXLENGTH=""13"" VALUE=""" & oRequest("RFC").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If StrComp(sAction,"Catalogs.asp",vbBinaryCompare) = 0 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP del empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CURP"" ID=""RFCTxt"" SIZE=""18"" MAXLENGTH=""18"" VALUE=""" & oRequest("CURP").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			End If
			If StrComp(sAction,"EmployeesDeleted",vbBinaryCompare) = 0 Then
				Response.Write "<TR>"
					Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""Status"" ID=""StatusID"" VALUE=""3"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			End If
			EmployeesDeleted
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresas:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyID"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write "<OPTION VALUE=""-1"">Ninguna</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyName", "(ParentID>-1) And (Active=1)", "CompanyName", oRequest("CompanyID").Item, "Ninguna;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			If StrComp(sAction,"Catalogs.asp",vbBinaryCompare) <> 0 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "(Active=1)", "EmployeeTypeName", oRequest("EmployeeTypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PositionTypeID"" ID=""PositionTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionTypes", "PositionTypeID", "PositionTypeName", "(Active=1) AND (PositionTypeID <> 6)", "PositionTypeName", oRequest("PositionTypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			If bFull Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Classificación:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""ClassificationIDTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & oRequest("ClassificationID").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "(Active=1)", "GroupGradeLevelName", oRequest("GroupGradeLevelID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""IntegrationIDTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & oRequest("IntegrationID").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelName", "(Active=1)", "LevelName", oRequest("LevelID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyName", "(Active=1)", "JourneyName", oRequest("JourneyID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftName", "(Active=1)", "ShiftName", oRequest("ShiftID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horas laboradas:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHoursTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & oRequest("WorkingHours").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""JobID"" ID=""JobIDTxt"" SIZE=""8"" MAXLENGTH=""8"" VALUE=""" & oRequest("JobID").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de Plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JobTypeID"" ID=""JobTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "JobTypeID", "JobTypeName", "(Active=1)", "JobTypeName", oRequest("JobTypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT><BR /><BR /></TD>"
				Response.Write "</TR>"
			Else
				If sAction = "Employees.asp" Then
					Response.Write "<TR>"
						Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><A HREF=""Employees.asp?Search=1""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Opciones avanzadas"" BORDER=""0"" HSPACE=""3"" ALIGN=""ABSMIDDLE"" />Opciones avanzadas</A></FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><BR /></TD>"
					Response.Write "</TR>"
				End If
			End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"""
				If Not bFull Then Response.Write " ALIGN=""RIGHT"""
				Response.Write "><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Empleados"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayEmployeesSearchForm = Err.number
End Function

Function DisplayEmployeeForms(oRequest, iSelectedTab, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the forms for the given employee
'Inputs:  oRequest, iSelectedTab, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeForms"
	Dim iIndex
	Dim lErrorNumber
	Dim sFontBegin
	Dim sFontEnd
	Dim bForm
    Dim sFilePath
    Dim sFileName
    Dim sFileNamepdf
    Dim sFileNamexml
    Dim bFileReady

	If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then
		sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
		sFontEnd = "</FONT>"
	Else
		sFontBegin = ""
		sFontEnd = ""
	End If
	If (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",5,6,", "," & iSelectedTab & ",", vbBinaryCompare) = 0) Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
'			Response.Write "<BR /><BR />"
'			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""30"" HEIGHT=""1"" />" & "<A HREF=""Employees.asp""><B>Consultar otro empleado</B></A>"
'			Response.Write "<BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP"" WIDTH=""400"">&nbsp;</TD>"
					Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><A HREF=""Employees.asp""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Terceros"" BORDER=""0"" /></A><BR /></TD>"
					Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""Employees.asp"" CLASS=""SpecialLink"">Otro empleado</A></B><BR /></FONT>"
					Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Consulte la información de un empleado diferente.</FONT></DIV></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write sFontBegin & "<B>Número de empleado: </B>" & CleanStringForHTML(aEmployeeComponent(S_NUMBER_EMPLOYEE)) & sFontEnd & "<BR />"
			Response.Write sFontBegin & "<B>Nombre: </B>" & CleanStringForHTML(aEmployeeComponent(S_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & sFontEnd & "<BR />"
			Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
			Response.Write sFontBegin & "<B>Estatus del empleado: </B>" & CleanStringForHTML(sNames) & sFontEnd & "<BR />"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR /><BR />"
			If iSelectedTab = 8 Then
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><A HREF=""http://192.168.2.134/DWWebClient/Login.aspx?DWSubSession=2423&v=1422"" target=""_blank""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Expediente Electrónico"" BORDER=""0"" /></A><BR /></TD>"
						Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""http://192.168.2.134/DWWebClient/Login.aspx?DWSubSession=2423&v=1422"" target=""_blank"" CLASS=""SpecialLink"">Expediente Electrónico</A></B><BR /></FONT>"
						Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Acceso al expediente electrónico.</FONT></DIV></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			End If
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "</FONT>"
	End If
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
		Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
			Select Case iSelectedTab
				Case 2
					lErrorNumber = DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), ",EmployeeManagement,", ",1,", -1, aEmployeeComponent, sErrorDescription)
				Case 3
					If Len(oRequest("PayrollID").Item) = 0 Then
						lErrorNumber = GetNameFromTable(oADODBConnection, "LastPayrollID", "-1", "", "", aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), sErrorDescription)
					Else
						aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = CDbl(oRequest("PayrollID").Item)
					End If
					Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Change"" ID=""ChangeHdn"" VALUE=""1"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""3"" />"
						Response.Write "Seleccione una nómina: <SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), "", sErrorDescription)
						Response.Write "</SELECT><BR />"
						Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Montos"" CLASS=""Buttons"" />"
					Response.Write "</FORM>"
					lErrorNumber = GetNameFromTable(oADODBConnection, "StatusEmployeesActive", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					If Len(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)) = 0 Then aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = -1
					If lErrorNumber = 0 Then
						If aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = -1 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("No existe una nómina abierta", "El administrador del módulo de nóminas debe crear una nueva nómina para registrar los conceptos de pago de los empleados.")
						ElseIf (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) Or (CInt(sNames) = 0) Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("El empleado no está activo", "El estatus del empleado no permite registrar los conceptos de pago en su nómina. Favor de cambiar su estatus a activo para continuar.")

							sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina del empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							lErrorNumber = DisplayEmployeeConceptsTable(oRequest, oADODBConnection, True, False, aEmployeeComponent, sErrorDescription)
						End If
					End If
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Case 4
					Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD WIDTH=""30%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" CLASS=""TableScrollDiv"">"
								Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceReview"" ID=""AbsenceReviewHdn"" VALUE=""1"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
									Response.Write "<B>Seleccione los datos para filtrar el historial:&nbsp;&nbsp;&nbsp;</B><BR />"
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar las incidencias del &nbsp;"
									Response.Write DisplayDateCombos(oRequest("FilterStartYear").Item, oRequest("FilterStartMonth").Item, oRequest("FilterStartDay").Item, "FilterStartYear", "FilterStartMonth", "FilterStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
									Response.Write "&nbsp;al&nbsp;"
									Response.Write DisplayDateCombos(oRequest("FilterEndYear").Item, oRequest("FilterEndMonth").Item, oRequest("FilterEndDay").Item, "FilterEndYear", "FilterEndMonth", "FilterEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
									Response.Write "<BR />"
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar las incidencias de: <SELECT NAME=""AbsenceID"" ID=""AbsenceIDCmb"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE=""-1"">Todas las claves</OPTION>"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(AbsenceID>0) And (AbsenceID<100) And (Active=1)", "AbsenceID", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "", sErrorDescription)
									Response.Write "</SELECT><BR />"
									Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons"" /><BR />"
									Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
								Response.Write "</FORM>"
								If lErrorNumber = 0 Then
									aAbsenceComponent(N_ACTIVE_ABSENCE) = 1
									lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aAbsenceComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									Response.Write "<BR />"
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
									bShowForm = True
								End If
							Response.Write "</DIV>"
						Response.Write "</FONT></TD>"
						'Response.Write "<TD>&nbsp;</TD>"
						'Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						'Response.Write "<TD>&nbsp;</TD>"
						'Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
						'	Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
						'		lErrorNumber = DisplayAbsenceForm(oRequest, oADODBConnection, GetASPFileName(""), "", "", aAbsenceComponent, sErrorDescription)
						'	Response.Write "</DIV>"
						'	If lErrorNumber <> 0 Then
						'		Response.Write "<BR />"
						'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						'		lErrorNumber = 0
						'		sErrorDescription = ""
						'	End If
						'Response.Write "</TD>"
					Response.Write "</TR></TABLE>"
				Case 5
					If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
						Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv"">"
							lErrorNumber = DisplayFormForHonoraryEmployee(oADODBConnection, False, aEmployeeComponent, sErrorDescription)
						Response.Write "</DIV>"
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Else
						Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv"">"
							'lErrorNumber = DisplayFormForEmployee(oADODBConnection, False, aEmployeeComponent, sErrorDescription)
							lErrorNumber = BuildReport1109(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
						Response.Write "</DIV>"
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					End If
				Case 6
					Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Change"" ID=""ChangeHdn"" VALUE=""1"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
						Response.Write "<B>Seleccione el historial que desea ver:&nbsp;&nbsp;&nbsp;</B>"
						Response.Write "<SELECT NAME=""ReportID"" ID=""ReportIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowHideReportFields(this.value)"">"
							Response.Write "<OPTION VALUE=""" & EMPLOYEE_HISTORY_LIST_REPORTS & """>Historial del empleado</OPTION>"
							'Response.Write "<OPTION VALUE=""" & EMPLOYEE_FORM_HISTORY_LIST_REPORTS & """>Historial de cambios del formato FM1</OPTION>"
							Response.Write "<OPTION VALUE=""" & EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS & """>Historial de pagos</OPTION>"
							Response.Write "<OPTION VALUE=""" & ISSSTE_1111_REPORTS & """>Historial de la plaza</OPTION>"
							Response.Write "<OPTION VALUE=""" & JOBS_LIST_REPORTS & """>Historial de ocupantes de la plaza</OPTION>"
							Response.Write "<OPTION VALUE=""" & EMPLOYEE_PAYROLL_REPORTS & """>Pagos por quincena</OPTION>"
						Response.Write "</SELECT><BR />"

						Response.Write "<DIV NAME=""PeriodDiv"" ID=""PeriodDiv"" STYLE=""display: none"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""LEFT"" />Mostrar el historial del &nbsp;"
							Response.Write DisplayDateCombos(oRequest("FilterStartYear").Item, oRequest("FilterStartMonth").Item, oRequest("FilterStartDay").Item, "FilterStartYear", "FilterStartMonth", "FilterStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
							Response.Write "&nbsp;al&nbsp;"
							Response.Write DisplayDateCombos(oRequest("FilterEndYear").Item, oRequest("FilterEndMonth").Item, oRequest("FilterEndDay").Item, "FilterEndYear", "FilterEndMonth", "FilterEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
							Response.Write "<BR />"
						Response.Write "</DIV>"

						Response.Write "<DIV NAME=""PayrollDiv"" ID=""PayrollDiv"" STYLE=""display: none"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""LEFT"" />Mostrar los pagos de &nbsp;"
							Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", oRequest("PayrollID").Item, "", sErrorDescription)
							Response.Write "</SELECT>"
							Response.Write "<BR />"
						Response.Write "</DIV>"

						Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons"" /><BR />"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"

						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function ShowHideReportFields(sValue) {" & vbNewLine
								Response.Write "var oForm = document.ReportFrm;" & vbNewLine

								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "switch (parseInt(sValue)) {" & vbNewLine
										Response.Write "case " & EMPLOYEE_HISTORY_LIST_REPORTS & ":" & vbNewLine
											Response.Write "ShowDisplay(document.all['PeriodDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['PayrollDiv']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case " & EMPLOYEE_PAYROLL_REPORTS & ":" & vbNewLine
											Response.Write "HideDisplay(document.all['PeriodDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['PayrollDiv']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "default:" & vbNewLine
											Response.Write "ShowDisplay(document.all['PeriodDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['PayrollDiv']);" & vbNewLine
											Response.Write "break;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "} // End of ShowHideReportFields" & vbNewLine

							Response.Write "SendURLValuesToForm('ReportID=" & oRequest("ReportID").Item & "', document.ReportFrm);" & vbNewLine
							Response.Write "ShowHideReportFields(document.ReportFrm.ReportID.value);" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					Response.Write "</FORM>"
					bForm = (((aLoginComponent(N_PROFILE_ID_LOGIN) <= 0) Or (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDeAntiguedades & ",", vbBinaryCompare) > 0)) And (CInt(oRequest("ReportID").Item) = EMPLOYEE_HISTORY_LIST_REPORTS) And (Len(oRequest("EmployeeDate").Item) > 0))
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD VALIGN=""TOP""><DIV NAME=""ReportDiv"" ID=""ReportDiv"""
							If bForm Then Response.Write " STYLE=""height: 350px; width:400px; overflow: auto;"""
						Response.Write ">"
							Select Case CInt(oRequest("ReportID").Item)
								Case EMPLOYEE_HISTORY_LIST_REPORTS
									lErrorNumber = DisplayEmployeeHistoryList(oRequest, oADODBConnection, False, True, aEmployeeComponent, sErrorDescription)
								Case EMPLOYEE_FORM_HISTORY_LIST_REPORTS
									lErrorNumber = DisplayEmployeeFormHistoryList(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
								Case EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS
									lErrorNumber = DisplayEmployeePaymentsHistoryList(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
								Case EMPLOYEE_PAYROLL_REPORTS
									lErrorNumber = DisplayEmployeePayroll(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
								Case ISSSTE_1111_REPORTS
									aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
									lErrorNumber = DisplayJobHistoryList(oRequest, oADODBConnection, False, False, aJobComponent, sErrorDescription)
								Case JOBS_LIST_REPORTS 
									aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
									lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
								Case Else
									Call DisplayInstructionsMessage("HISTORIAL DEL EMPLEADO", "Seleccione el historial que desea ver y, en caso de ser necesario, indique los periodos para acotar la información.")
							End Select
						Response.Write "</DIV></TD>"
						If bForm Then
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
								lErrorNumber = ShowEmployeeHistoryListForm(oRequest, oADODBConnection, GetASPFileName(""), aEmployeeComponent, sErrorDescription)
							Response.Write "</TD>"
						End If
					Response.Write "</TR></TABLE>"
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Case 7
					Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Change"" ID=""ChangeHdn"" VALUE=""1"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CancelPayment"" ID=""CancelPaymentHdn"" VALUE=""1"" />"
						Response.Write "<B>Seleccione el reporte que desea ver:&nbsp;&nbsp;&nbsp;</B>"
						Response.Write "<SELECT NAME=""ReportID"" ID=""ReportIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowHideReportFields(this.value)"">"
							Response.Write "<OPTION VALUE=""" & ISSSTE_1116_REPORTS & """>Antigüedad</OPTION>"
							Response.Write "<OPTION VALUE=""" & ISSSTE_1208_REPORTS & """>Constancia de descuento</OPTION>"
							Response.Write "<OPTION VALUE=""" & ISSSTE_1003_REPORTS & """>Listado de firmas</OPTION>"
							Response.Write "<OPTION VALUE=""" & ISSSTE_1338_REPORTS & """>Revisión de nóminas</OPTION>"
							Response.Write "<OPTION VALUE=""-100"">Cheques cancelados</OPTION>"
							'Response.Write "<OPTION VALUE=""" & ISSSTE_1002_REPORTS & """>Revisión de diferencias</OPTION>"
						Response.Write "</SELECT><BR />"

						Response.Write "<DIV NAME=""PayrollDiv"" ID=""PayrollDiv"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""20"" ALIGN=""LEFT"" />Mostrar los pagos de &nbsp;"
							Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", oRequest("PayrollID").Item, "", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</DIV>"

						Response.Write "<DIV NAME=""ConceptsDiv"" ID=""ConceptsDiv"" STYLE=""display: none"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""72"" ALIGN=""LEFT"" />"
							Response.Write "<B>Conceptos de pago:</B><BR />"
							Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(EndDate=30000000)", "ConceptShortName, ConceptName", "", "", sErrorDescription)
							Response.Write "</SELECT><BR /><BR />"
							Response.Write "<B>Periodo:</B>"
							Response.Write "&nbsp;Entre&nbsp;"
							Response.Write DisplayDateCombos(CInt(oRequest("StartYear").Item), CInt(oRequest("StartMonth").Item), CInt(oRequest("StartDay").Item), "StartYear", "StartMonth", "StartDay", 2008, Year(Date()), True, False)
							Response.Write "&nbsp;y el&nbsp;"
							Response.Write DisplayDateCombos(CInt(oRequest("EndYear").Item), CInt(oRequest("EndMonth").Item), CInt(oRequest("EndDay").Item), "EndYear", "EndMonth", "EndDay", 2008, Year(Date()), True, False)
						Response.Write "</DIV>"

						Response.Write "<DIV NAME=""AntiquityDiv"" ID=""AntiquityDiv"" STYLE=""display: none"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""20"" ALIGN=""LEFT"" />"
							Response.Write "<B>Antigüedad hasta el día: </B>"
							If (Len(oRequest("EmployeeYear").Item) > 0) And (Len(oRequest("EmployeeMonth").Item) > 0) And (Len(oRequest("EmployeeDay").Item) > 0) Then
								Response.Write DisplayDateCombos(CInt(oRequest("EmployeeYear").Item), CInt(oRequest("EmployeeMonth").Item), CInt(oRequest("EmployeeDay").Item), "EmployeeYear", "EmployeeMonth", "EmployeeDay", N_FORM_START_YEAR, Year(Date()) + 1, True, False)
							Else
								Response.Write DisplayDateCombos(Year(Date()), Month(Date()), Day(Date()), "EmployeeYear", "EmployeeMonth", "EmployeeDay", N_FORM_START_YEAR, Year(Date()) + 1, True, False)
							End If
						Response.Write "</DIV>"

						Response.Write "<DIV NAME=""Note1002Div"" ID=""Note1002Div"">"
							Response.Write "<B>Nota: </B>Esta nómina será comparada contra la nómina anterior.<BR /><BR />"
						Response.Write "</DIV>"

						Response.Write "<BR /><BR />"
						Response.Write "<DIV NAME=""ContinueDiv"" ID=""ContinueDiv""><INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons"" /></DIV>"
						Response.Write "<DIV NAME=""Continue2Div"" ID=""Continue2Div"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""20"" ALIGN=""LEFT"" />"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fecha de emisión de la nómina:&nbsp;</FONT>"
							Response.Write DisplayDateCombos(CInt(oRequest("PayrollIssueYear").Item), CInt(oRequest("PayrollIssueMonth").Item), CInt(oRequest("PayrollIssueDay").Item), "PayrollIssueYear", "PayrollIssueMonth", "PayrollIssueDay", N_FORM_START_YEAR, Year(Date()), True, False)
							Response.Write "<BR /><BR />"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("PayrollIssueYear").Item) = 0 Then Response.Write "document.ReportFrm.PayrollIssueYear.value = " & Year(Date()) & ";" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
							Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Ver Reporte"" CLASS=""Buttons"" onClick=""window.location.href = 'Employees.asp?Change=1&Tab=7&ReportID=' + document.ReportFrm.ReportID.value + '&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&PayrollID=' + document.ReportFrm.PayrollID.value + '&PayrollIssueYear=' + document.ReportFrm.PayrollIssueYear.value + '&PayrollIssueMonth=' + document.ReportFrm.PayrollIssueMonth.value + '&PayrollIssueDay=' + document.ReportFrm.PayrollIssueDay.value + '';"" />"
						Response.Write "</DIV>"
						Response.Write "<DIV NAME=""Continue3Div"" ID=""Continue3Div""><INPUT TYPE=""BUTTON"" VALUE=""Ver Reporte"" CLASS=""Buttons"" onClick=""window.location.href = 'Employees.asp?Change=1&Tab=7&ReportID=1116&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeYear=' + document.ReportFrm.EmployeeYear.value + '&EmployeeMonth=' + document.ReportFrm.EmployeeMonth.value + '&EmployeeDay=' + document.ReportFrm.EmployeeDay.value;"" /></DIV>"
						Response.Write "<BR />"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"

						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function ShowHideReportFields(sValue) {" & vbNewLine
								Response.Write "var oForm = document.ReportFrm;" & vbNewLine

								Response.Write "ShowDisplay(document.all['PayrollDiv']);" & vbNewLine
								Response.Write "ShowDisplay(document.all['ContinueDiv']);" & vbNewLine
								Response.Write "HideDisplay(document.all['Note1002Div']);" & vbNewLine
								Response.Write "HideDisplay(document.all['ConceptsDiv']);" & vbNewLine
								Response.Write "HideDisplay(document.all['AntiquityDiv']);" & vbNewLine
								Response.Write "HideDisplay(document.all['Continue2Div']);" & vbNewLine
								Response.Write "HideDisplay(document.all['Continue3Div']);" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "switch (sValue) {" & vbNewLine
										Response.Write "case '" & ISSSTE_1002_REPORTS & "':" & vbNewLine
											Response.Write "ShowDisplay(document.all['Note1002Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '" & ISSSTE_1003_REPORTS & "':" & vbNewLine
											Response.Write "HideDisplay(document.all['ContinueDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['Continue2Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '" & ISSSTE_1116_REPORTS & "':" & vbNewLine
											Response.Write "HideDisplay(document.all['PayrollDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['ContinueDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['AntiquityDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['Continue3Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '" & ISSSTE_1208_REPORTS & "':" & vbNewLine
											Response.Write "HideDisplay(document.all['PayrollDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['ConceptsDiv']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '" & ISSSTE_1338_REPORTS & "':" & vbNewLine
											Response.Write "HideDisplay(document.all['ContinueDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['Continue2Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '-100':" & vbNewLine
											Response.Write "HideDisplay(document.all['PayrollDiv']);" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "} // End of ShowHideReportFields" & vbNewLine

							Response.Write "SendURLValuesToForm('ReportID=" & oRequest("ReportID").Item & "', document.ReportFrm);" & vbNewLine
							Response.Write "ShowHideReportFields(document.ReportFrm.ReportID.value);" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					Response.Write "</FORM>"
					Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv"">"
						Select Case CInt(oRequest("ReportID").Item)
							Case -100
								lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 0, False, sErrorDescription)
								'lErrorNumber = DisplayPaymentsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, False, aPaymentComponent, sErrorDescription)
							Case ISSSTE_1002_REPORTS
								lErrorNumber = BuildReport1002(oRequest, oADODBConnection, False, sErrorDescription)
							Case ISSSTE_1003_REPORTS
								lErrorNumber = BuildReports1003(oRequest, oADODBConnection, False, sErrorDescription)
							Case ISSSTE_1116_REPORTS
								lErrorNumber = BuildReport1116(oRequest, oADODBConnection, False, Null, sErrorDescription)
							Case ISSSTE_1208_REPORTS
								lErrorNumber = BuildReport1208(oRequest, oADODBConnection, False, sErrorDescription)
							Case ISSSTE_1338_REPORTS
								lErrorNumber = BuildReports1003(oRequest, oADODBConnection, True, sErrorDescription)
							Case Else
								Call DisplayInstructionsMessage("REPORTES SOBRE EL EMPLEADO", "Seleccione el reporte que desea ver y, en caso de ser necesario, indique la nómina que desea utilizar.")
						End Select
					Response.Write "</DIV>"
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Case 8
					Response.Write "<IFRAME SRC=""BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&Tab=8"" NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
					If Len(SECOND_PHYSICAL_PATH) > 0 Then
						Response.Write "</FONT></TD>"
						Response.Write "<TD WIDTH=""1"">&nbsp;&nbsp;&nbsp;</TD>"
						Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD WIDTH=""1"">&nbsp;</TD>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
							If InStr(1, SECOND_PHYSICAL_PATH, "http", vbBinaryCompare) = 1 Then
								Response.Write "<IFRAME SRC=""" & Replace(Replace(Replace(Replace(SECOND_PHYSICAL_PATH, "<EMPLOYEE_ID />", aEmployeeComponent(S_NUMBER_EMPLOYEE)), "<EMPLOYEE_NAME />", aEmployeeComponent(S_NAME_EMPLOYEE)), "<EMPLOYEE_LAST_NAME />", aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "<EMPLOYEE_LAST_NAME2 />", aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
							Else
								Response.Write "<IFRAME SRC=""BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&SecondFolder=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
							End If
					End If
				Case 9
                    Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
		            Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"					
	                If (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",3,5,6,7,", "," & iSelectedTab & ",", vbBinaryCompare) = 0) Then
			            Response.Write "</FONT></TD>"
			            Response.Write "<TD WIDTH=""1"">&nbsp;&nbsp;&nbsp;</TD>"
			            Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			            Response.Write "<TD WIDTH=""1"">&nbsp;</TD>"
			            Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"

		            End If
         
                    Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE="""  & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
                        Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nómina:<BR /></FONT>"
                        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""DisplayPayrollFilters(this.value)"">"
				        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName",  "Where (PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", "", "", sErrorDescription)
	    		        Response.Write "</SELECT><BR /><BR />"			
						Response.Write "<BR /><BR />"
						Response.Write "<DIV NAME=""FactDiv"" ID=""FactDiv""><INPUT TYPE=""SUBMIT"" VALUE=""Ver Factura"" CLASS=""Buttons"" /></DIV>"
                        If Len(oRequest("PayrollID").Item)>0 Then
                           Response.Write "<BR /><BR />"
                            Response.Write "<DIV NAME=""Continue2Div"" ID=""Continue2Div"">"
							Response.Write "<BR />"
						    Response.Write "</DIV>"
                            sFileName = "Facturas\" & oRequest("PayrollID").Item & "\" & oRequest("PayrollID").Item & "_" & aEmployeeComponent(N_ID_EMPLOYEE) &".zip"
                            sFileName = Server.MapPath(sFileName)
                            If Not FileExists(sFileName, sErrorDescription) Then
                                sFileNamepdf = "Facturas\" & oRequest("PayrollID").Item & "\" & oRequest("PayrollID").Item & "_" & aEmployeeComponent(N_ID_EMPLOYEE) &".pdf"
                                sFileNamepdf = Server.MapPath(sFileNamepdf)
                                sFileNamexml =  "Facturas\" & oRequest("PayrollID").Item & "\" & oRequest("PayrollID").Item & "_" & aEmployeeComponent(N_ID_EMPLOYEE) &".xml"
                                sFileNamexml = Server.MapPath(sFileNamexml)
                                sErrorDescription = "La factura no se encuentra disponible"
                                If (FileExists(sFileNamepdf, sErrorDescription)) OR (FileExists(sFileNamexml, sErrorDescription)) Then
                                    sFilePath = Server.MapPath("Facturas\" & oRequest("PayrollID").Item)& "\"& oRequest("PayrollID").Item & "_" & aEmployeeComponent(N_ID_EMPLOYEE) & "\"
                                    Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
    			                    Response.Flush()
                                    lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
                                    If lErrorNumber = 0 Then
                                        lErrorNumber = CopyFile(sFileNamepdf, sFilePath, sErrorDescription)
                                        lErrorNumber = CopyFile(sFileNamexml, sFilePath, sErrorDescription)
                                        If lErrorNumber = 0 Then
                                            lErrorNumber = ZipFolder(sFilePath, sFileName, sErrorDescription)
                                            lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
                                        End If    
                                    End IF
                                 Else 
                                    sErrorDescription = "La factura no se encuentra disponible"
                                    Response.Write "<BR />"
						            Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						            lErrorNumber = 0
						            sErrorDescription = ""
                                 End IF    
                            Else
                                Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"                           
                            End If                         
                            oEndDate = Now()
                        End IF
						
					Response.Write "</FORM>"
				
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If	
		Response.Write "</FONT></TD>"
	Response.Write "</TR></TABLE>"
                        
					'If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
					'	Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv"">"
					'		lErrorNumber = DisplayFormForHonoraryEmployee(oADODBConnection, False, aEmployeeComponent, sErrorDescription)
					'	Response.Write "</DIV>"
					'	If lErrorNumber <> 0 Then
					'		Response.Write "<BR />"
					'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					'		lErrorNumber = 0
					'		sErrorDescription = ""
					'	End If
					'Else
					'	Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv"">"
					'		'lErrorNumber = DisplayFormForEmployee(oADODBConnection, False, aEmployeeComponent, sErrorDescription)
					'		lErrorNumber = BuildReport1158(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
					'	Response.Write "</DIV>"
					'	If lErrorNumber <> 0 Then
					'		Response.Write "<BR />"
					'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					'		lErrorNumber = 0
					'		sErrorDescription = ""
					'	End If
					'End If
				Case Else
					If Len(oRequest.Item("ShowInfo")) > 0 Then
						lErrorNumber = DisplayEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), ",EmployeeManagement,", ",1,", -1, aEmployeeComponent, sErrorDescription)
						Response.Write "<DIV STYLE=""display: none"">"
							lErrorNumber = DisplayEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						Response.Write "</DIV>"
					End If
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
			End Select
		If (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",3,5,6,7,", "," & iSelectedTab & ",", vbBinaryCompare) = 0) Then
			Response.Write "</FONT></TD>"
			Response.Write "<TD WIDTH=""1"">&nbsp;&nbsp;&nbsp;</TD>"
			Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""1"">&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
		End If
			Select Case iSelectedTab
				Case 2
					If aEmployeeComponent(N_JOB_ID_EMPLOYEE) > -1 Then
						aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
						lErrorNumber = DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia de la ocupación de la plaza:</B></FONT>"
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
							lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
						End If
					Else
						If Not B_ISSSTE Then	
							Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckRadioSelection(this.JobID)"">"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeeJob"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
								lErrorNumber = DisplayFreeJobsTable(oRequest, oADODBConnection, DISPLAY_RADIO_BUTTONS, False, aJobComponent, sErrorDescription)
								Response.Write "<BR /><INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Asignar Plaza"" CLASS=""Buttons"" />"
							Response.Write "</FORM>"
						Else
							Response.Write "El empleado no tiene asignada ninguna plaza"
						End If
					End If
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Error en la información del empleado", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Case 3
					'If (aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) > -1) And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1) And (CInt(sNames) = 1) Then
					'	lErrorNumber = DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "", "", aEmployeeComponent, sErrorDescription)
					'	If lErrorNumber <> 0 Then
					'		Response.Write "<BR />"
					'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					'		lErrorNumber = 0
					'		sErrorDescription = ""
					'	End If
					'End If
				Case 4
					'lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, False, aAbsenceComponent, sErrorDescription)
					'If lErrorNumber <> 0 Then
					'	Response.Write "<BR />"
					'	Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					'	lErrorNumber = 0
					'	sErrorDescription = ""
					'End If
				Case 5
				Case 6
				Case 7
				Case 8
                Case 9   
				Case Else
					If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
						Response.Write "<B>Hijos</B><BR />"
						lErrorNumber = DisplayEmployeeChildrenTable(oRequest, oADODBConnection, "Employees", DISPLAY_NOTHING, True, False, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
						'If Len(oRequest("ShowInfo").Item) = 0 Then
						'	lErrorNumber = DisplayEmployeeChildForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesChildren", "", aEmployeeComponent, sErrorDescription)
						'	If lErrorNumber <> 0 Then
						'		Response.Write "<BR />"
						'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						'		lErrorNumber = 0
						'		sErrorDescription = ""
						'	End If
						'End If
						Response.Write "<BR /><BR />"
						'Response.Write "<B>Beneficiarios</B><BR />"
						'lErrorNumber = DisplayEmployeeBeneficiariesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aEmployeeComponent, sErrorDescription)
						'If lErrorNumber <> 0 Then
						'	Response.Write "<BR />"
						'	Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						'	lErrorNumber = 0
						'	sErrorDescription = ""
						'End If
						'If Len(oRequest("ShowInfo").Item) = 0 Then
						'	lErrorNumber = DisplayEmployeeBeneficiaryForm(oRequest, oADODBConnection, GetASPFileName(""), "", aEmployeeComponent, sErrorDescription)
						'	If lErrorNumber <> 0 Then
						'		Response.Write "<BR />"
						'		Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						'		lErrorNumber = 0
						'		sErrorDescription = ""
						'	End If
						'End If
					End If
			End Select
		Response.Write "</FONT></TD>"
	Response.Write "</TR></TABLE>"

	DisplayEmployeeForms = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesTabs(oRequest, bError, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the employees HTML forms
'Inputs:  oRequest, bError
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesTabs"
	Dim asTitles
	Dim iIndex
	Dim sAction
	Dim lErrorNumber

	If aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> -2 Then
		If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
			asTitles = Split(",Información del empleado,Plaza,Conceptos de pago,Incidencias,Formato honorarios,Historiales,Reportes,Expediente Electrónico,Factura Electrónica", ",")
		Else
			asTitles = Split(",Información del empleado,Plaza,Conceptos de pago,Incidencias,Formato FM1,Historiales,Reportes,Expediente Electrónico,Factura Electrónica", ",")
		End If
		If B_ISSSTE Then
		Else
			If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_JOBS_PERMISSIONS) <> N_JOBS_PERMISSIONS Then asTitles(2) = ""
			If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_EMPLOYEE_PAYROLL_PERMISSIONS) <> N_EMPLOYEE_PAYROLL_PERMISSIONS Then
				asTitles(3) = ""
				asTitles(4) = ""
			End If
			If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REPORTS_PERMISSIONS) <> N_REPORTS_PERMISSIONS Then
				asTitles(6) = ""
				asTitles(7) = ""
			End If
		End If
	ElseIf (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = -2) And (Len(oRequest("New").Item) > 0) Then
		asTitles = Split(",Asignación de número de empleado", ",")
	ElseIf aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = -2 Then
		asTitles = Split(",Información del empleado", ",")
	Else
		asTitles = Split(",Información del empleado,Plaza,,,Formato FM1,,", ",")
	End If
	If (Len(oRequest("New").Item) > 0) Or (bError And (StrComp(oRequest("Action").Item, "Employees", vbBinaryCompare) = 0) And (Len(oRequest("Add").Item) > 0)) Then
		Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""5"" NAME=""TabContents1LfDiv"" ID=""TabContents1LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ BACKGROUND=""Images/TbBg.gif"" WIDTH=""130"" ALIGN=""CENTER"" NAME=""TabContents1Div"" ID=""TabContents1Div""><NOBR><FONT FACE=""Arial"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """ SIZE=""2"" CLASS=""TabLink"">"
			Response.Write "<B>&nbsp;&nbsp;&nbsp;" & asTitles(1) & "&nbsp;&nbsp;&nbsp;</B></FONT></NOBR></TD>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""5"" NAME=""TabContents1RgDiv"" ID=""TabContents1RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
			Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""21"" /></TD>"
		Response.Write "</TR></TABLE><BR />"
	Else
		sAction = "ShowInfo"
		If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sAction = "Change"
		Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			For iIndex = 1 To UBound(asTitles)
				If Len(asTitles(iIndex)) > 0 Then
					Response.Write "<TD BGCOLOR=""#"
						If iSelectedTab = iIndex Then
							Response.Write S_MAIN_COLOR_FOR_GUI
						Else
							Response.Write "CCCCCC"
						End If
					Response.Write """ WIDTH=""5"" NAME=""TabContents" & iIndex & "LfDiv"" ID=""TabContents" & iIndex & "LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
					Response.Write "<TD BGCOLOR=""#"
						If iSelectedTab = iIndex Then
							Response.Write S_MAIN_COLOR_FOR_GUI
						Else
							Response.Write "CCCCCC"
						End If
					Response.Write """ BACKGROUND=""Images/TbBg.gif"" WIDTH=""130"" ALIGN=""CENTER"" NAME=""TabContents" & iIndex & "Div"" ID=""TabContents" & iIndex & "Div""><NOBR><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<A HREF=""Employees.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&" & sAction & "=1&Tab=" & iIndex & """ CLASS=""TabLink""><DIV NAME=""TabText" & iIndex & "Div"" ID=""TabText" & iIndex & "Div"" STYLE=""color: #"
						If iSelectedTab = iIndex Then
							Response.Write S_MENU_LINK_FOR_GUI
						Else
							Response.Write "000000"
						End If
					Response.Write ";""><B>&nbsp;&nbsp;&nbsp;" & asTitles(iIndex) & "&nbsp;&nbsp;&nbsp;</B></DIV></A></FONT></NOBR></TD>"
					Response.Write "<TD BGCOLOR=""#"
						If iSelectedTab = iIndex Then
							Response.Write S_MAIN_COLOR_FOR_GUI
						Else
							Response.Write "CCCCCC"
						End If
					Response.Write """ WIDTH=""5"" NAME=""TabContents" & iIndex & "RgDiv"" ID=""TabContents" & iIndex & "RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
				End If
			Next
			Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""21"" HEIGHT=""21"" /></TD>"
		Response.Write "</TR></TABLE>"
	End If

	DisplayEmployeesTabs = lErrorNumber
	Err.Clear
End Function
%>