<%
If Datediff("d", Now(), Application.Contents("SIAP_DaemonDay")) < 0 Then
	Application.Contents("SIAP_DaemonDay") = Now()
	Application.Contents("SIAP_DaemonStatus") = 0
End If

Function UpdateCurrenciesHistory(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To initialize the currencies values
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateCurrenciesHistory"
	Dim oTempDate
	Dim sTempDate
	Dim oRecordset
	Dim lErrorNumber

	If (CLng(Application.Contents("SIAP_DaemonStatus")) And N_UPDATE_CURRENCIES_DAEMON) <> N_UPDATE_CURRENCIES_DAEMON Then 
		Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) Or N_UPDATE_CURRENCIES_DAEMON
		sErrorDescription = "No se pudo obtener la fecha de la última actualización al historial de las monedas."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(CurrencyDate) From CurrenciesHistoryList", "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			oTempDate = DateSerial(Left(CStr(oRecordset.Fields(0).Value), 4), Mid(CStr(oRecordset.Fields(0).Value), 5, 2), Right(CStr(oRecordset.Fields(0).Value), 2))
			oTempDate = DateAdd("d", 1, oTempDate)
			oRecordset.Close
			sErrorDescription = "No se pudieron obtener los valores de las monedas registradas en el sistema."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyID, CurrencyValue From Currencies Order By CurrencyID", "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Do While DateDiff("d", oTempDate, Date()) >= 0
				sTempDate = Year(oTempDate) & Right(("0" & Month(oTempDate)), Len("00")) & Right(("0" & Day(oTempDate)), Len("00"))
				Do While Not oRecordset.EOF
					sErrorDescription = "No se pudo modificar el historial de las monedas."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into CurrenciesHistoryList (CurrencyID, CurrencyDate, CurrencyValue) Values (" & CStr(oRecordset.Fields("CurrencyID").Value) & ", " & sTempDate & ", " & CStr(oRecordset.Fields("CurrencyValue").Value) & ")", "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oTempDate = DateAdd("d", 1, oTempDate)
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
		If (lErrorNumber <> 0) Then
			Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) - N_UPDATE_CURRENCIES_DAEMON
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	UpdateCurrenciesHistory = lErrorNumber
	Err.Clear
End Function

Function AddEmployeesConceptC5Dm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To Select employees with antiquities for add
'         concept C5
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeesConceptC5Dm"
	Dim oTempDate
	Dim sTempDate
	Dim oRecordset
	Dim lErrorNumber
	Dim lDate

	lDate = AddDaysToSerialDate(Left(GetSerialNumberForDate(""), Len("00000000")), -9125)
	If (CLng(Application.Contents("SIAP_DaemonStatus")) And N_ADD_EMPLOYEES_CONCEPT_C5) <> N_ADD_EMPLOYEES_CONCEPT_C5 Then
		Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) Or N_UPDATE_EMPLOYEES_STATUS_DAEMON
		sErrorDescription = "No se pudo obtener el listado de empleados con antiguedad de 25 años."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (StartDate=" & lDate & ") And (StatusID IN (0,1))", "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					'Codigo para actualizar estatus de empleados cuando inicia la suspensión
					aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
					aEmployeeComponent(N_JOB_ID_EMPLOYEE) = CLng(oRecordset.Fields("JobID").Value)
					lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = CalculateAmountDiferenceForAntiquityConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 5
							aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = Left(GetSerialNumberForDate(""), Len("00000000"))
							lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			End If
		End If
	End If
	If (lErrorNumber <> 0) Then
		Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) - N_UPDATE_EMPLOYEES_STATUS_DAEMON
	End If
	oRecordset.Close
	Set oRecordset = Nothing
	AddEmployeesConceptC5Dm = lErrorNumber
	Err.Clear
End Function

Function UpdateEmployeesReasonDm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To update the current status of the employees
'         by reason at the end date
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "UpdateEmployeesReasonDm"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery
	Dim sQuery2
	Dim sUpdate
	Dim sInsert
	Dim lCurrentDate
	Dim lLoop
	Dim sResponseParameters
	Dim lPayrollID
	Dim lReasonID

	lErrorNumber = 0
	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	If (CLng(Application.Contents("SIAP_DaemonStatus")) And N_UPDATE_EMPLOYEES_REASON_DAEMON) <> N_UPDATE_EMPLOYEES_REASON_DAEMON Then 	
		Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) Or N_UPDATE_EMPLOYEES_REASON_DAEMON
		'Se obtienen la última quincena abierta
		lPayrollID = -1
		sQuery = "Select PayrollID From Payrolls Where (IsClosed<>1) And (PayrollTypeID <> 0) Order By PayrollID Desc"
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then lPayrollID = oRecordset.Fields("PayrollID").Value
			oRecordset.Close
		End If

		'Baja de Riesgo Profesional, Turno Extra y Percepción Adicional
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sUpdate = "Update EmployeesConceptsLKP Set Active = 0 Where (ConceptID In (4,7,8)) And (EndDate < " & lCurrentDate & ") And (Active = 1)"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sUpdate, "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber <> 0 Then
					sErrorDescription = "Los conceptos 04 (Riesgos Profesionales), 07 (Turno Extra) y 08 (Percepción Extra) no pudieron cerrarse"
				End If
			End If
		Else
			sErrorDescription = "No se pudieron obtener los conceptos para cerrar"
		End If
		oRecordset.Close

		sQuery = "Select EmployeeID, JobID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, " & _
				"EmployeeTypeID, RFC, CURP, SocialSecurityNumber, BirthDate, GenderID, MaritalStatusID, " & _
				"EmployeeEmail, ShiftID From Employees Where"
		sResponseParameters = ""

		For lLoop = 1 To 18
			Select Case lLoop
				Case 1:	'Empleados por honorarios
						sCondition = " (StatusID = 0) And (Active = 1) And (EmployeeTypeID = 7) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 0) And (Active = 1) And (EmployeeTypeID = 7) And (ReasonID = 14) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 14
				Case 2: 'Empleados en interinato
						sCondition = " (StatusID = 1) And (Active = 1) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 1) And (Active = 1) And (ReasonID = 13) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 13
				Case 3: 'Empleados en puesto de confianza dentro del instituto
						sCondition = " (StatusID = 0) And (Active = 1) And (PositionTypeID = 2) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 0) And (Active = 1) And (PositionTypeID = 2) And (ReasonID = 68) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 68
'				Case 4: 'Cambio de plaza misma adscripción
'						sCondition = ""
'				Case 5: 'Cambio de adscripción con plaza
'						sCondition = ""
'				Case 6: 'Cambio de adscripción sin plaza
'						sCondition = ""
'				Case 7: 'Permuta de plazas
'						sCondition = ""
				Case 8: 'Licencia sin goce de sueldo por asuntos particulares
						sCondition = " (StatusID = 78) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 78) And (ReasonID = 45) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 45
				Case 9: 'Licencia sin goce de sueldo por comisión sindical
						sCondition = " (StatusID = 82) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 82) And (ReasonID = 38) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 38
				Case 10: 'Licencia sin goce de sueldo por otorgamiento de beca
						sCondition = " (StatusID = 90) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 90) And (ReasonID = 48) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 48
				Case 11: 'Licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del insituto
						sCondition = " (StatusID = 94) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 94) And (ReasonID = 43) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 43
				Case 12: 'Licencia sin goce de sueldo por ocupar puesto de confianza dentro del instituto
						sCondition = " (StatusID = 98) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 98) And (ReasonID = 46) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 46
				Case 13: 'Licencia sin goce de sueldo por práctica de servicio social
						sCondition = " (StatusID = 102) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 102) And (ReasonID = 39) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonId = 39
				Case 14: 'Prórroga de licencia sin goce de sueldo por comisión sindical
						sCondition = " (StatusID = 106) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 106) And (ReasonID = 47) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 47
				Case 15: 'Prórroga de licencia sin goce de sueldo por otorgamiento de beca
						sCondition = " (StatusID = 110) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 110) And (ReasonID = 40) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 40
				Case 16: 'Prórroga de licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del instituto
						sCondition = " (StatusID = 114) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 114) And (ReasonID = 44) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 44
				Case 17: 'Prórroga de licencia sin goce de sueldo por ocupar puesto de confianza dentro del instituto
						sCondition = " (StatusID = 118) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 118) And (ReasonID = 41) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 41
				Case 18: 'Prórroga de licencia sin goce de sueldo por asuntos particulares
						sCondition = " (StatusID = 140) And (EmployeeID In (Select Distinct EmployeeID From EmployeesHistoryList Where (StatusID = 140) And (ReasonID = 37) And (EndDate < " & lCurrentDate & "))) Order By EmployeeID"
						lReasonID = 37
			End Select

			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition, "DaemonLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sResponseParameters = "EmployeeID=" & oRecordset.Fields("EmployeeID").Value & "&ReasonID=" & lReasonID & _
					"&SaveEmployeesMovements=1&JobID=" & oRecordset.Fields("JobID").Value & "&EmployeeNumber=" & oRecordset.Fields("EmployeeNumber").Value & _
					"&EmployeeName=" & oRecordset.Fields("EmployeeName").Value & "&EmployeeLastName=" & oRecordset.Fields("EmployeeLastName").Value & _
					"&EmployeeLastName2=" & oRecordset.Fields("EmployeeLastName2").Value & "&EmployeeTypeID=" & oRecordset.Fields("EmployeeTypeID").Value & _
					"&RFC=" & oRecordset.Fields("RFC").Value & "&CURP=" & oRecordset.Fields("CURP").Value & "&EmployeeDay=" & Mid(lCurrentDate,7,2) & _
					"&EmployeeMonth=" & Mid(lCurrentDate,5,2) & "&EmployeeYear=" & Mid(lCurrentDate,1,4) & "&IsBatch=1"
				If lReasonID = 14 Then
					sResponseParameters = sResponseParameters & "&DropReasonName=Baja+por+renuncia"
				End If
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						Response.Redirect "UploadInfo.asp?" & sResponseParameters
					Loop
				End If
			End If
		Next
	End If

	If (lErrorNumber <> 0) Then
		Application.Contents("SIAP_DaemonStatus") = CLng(Application.Contents("SIAP_DaemonStatus")) - N_UPDATE_EMPLOYEES_REASON_DAEMON
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	UpdateEmployeesReasonDm = lErrorNumber
	Err.Clear
End Function
%>