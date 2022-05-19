<%
Const N_EMPLOYEE_ID_REVISION = 0
Const N_CONCEPT_ID_REVISION = 1
Const N_START_DATE_REVISION = 2
Const N_END_DATE_REVISION = 3
Const D_CONCEPT_AMOUNT_REVISION = 4
Const N_DATE_PAYROLL_REVISION = 5
Const N_MODIFY_DATE_REVISION = 6
Const N_USER_ID_REVISION = 7
Const N_ACTIVE_REVISION = 8
Const S_COMMENTS_REVISION = 9
Const B_IS_DUPLICATED_REVISION = 10
Const B_COMPONENT_INITIALIZED_REVISION = 11
Const S_QUERY_CONDITION_REVISION = 12

Const N_REVISION_COMPONENT_SIZE = 12

Dim aPayrollRevisionComponent()
Redim aPayrollRevisionComponent(N_REVISION_COMPONENT_SIZE)

Function InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
'************************************************************
'Purpose: To initialize the empty elements of the PayrollRevision
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aConceptComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePayrollRevisionComponent"
	Redim Preserve aPayrollRevisionComponent(N_CONCEPT_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = CLng(oRequest("EmployeeID").Item)
		Else
			aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_CONCEPT_ID_REVISION)) Then
		If Len(oRequest("ConceptID").Item) > 0 Then
			aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) = CLng(oRequest("ConceptID").Item)
		Else
			aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_START_DATE_REVISION)) Then
		If Len(oRequest("StartPayrollID").Item) > 0 Then
			aPayrollRevisionComponent(N_START_DATE_REVISION) = CLng(oRequest("StartPayrollID").Item)
		Else
			aPayrollRevisionComponent(N_START_DATE_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_END_DATE_REVISION)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aPayrollRevisionComponent(N_END_DATE_REVISION) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aPayrollRevisionComponent(N_END_DATE_REVISION) = CLng(oRequest("EndDate").Item)
		Else
			aPayrollRevisionComponent(N_END_DATE_REVISION) = 30000000
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION)) Then
		If Len(oRequest("ConceptAmount").Item) > 0 Then
			aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = CDbl(oRequest("ConceptAmount").Item)
		Else
			aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) = 0
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION)) Then
		If Len(oRequest("PayrollDateYear").Item) > 0 Then
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = CInt(oRequest("PayrollDateYear").Item) & Right(("0" & oRequest("PayrollDateMonth").Item), Len("00")) & Right(("0" & oRequest("PayrollDateDay").Item), Len("00"))
		ElseIf Len(oRequest("PayrollDate").Item) > 0 Then
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = CLng(oRequest("PayrollDate").Item)
		Else
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_MODIFY_DATE_REVISION)) Then
		If Len(oRequest("ModifyDate").Item) > 0 Then
			aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) = CInt(oRequest("ModifyDateYear").Item) & Right(("0" & oRequest("ModifyDateMonth").Item), Len("00")) & Right(("0" & oRequest("ModifyDateDay").Item), Len("00"))
		ElseIf Len(oRequest("ModifyDate").Item) > 0 Then
			aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) = CLng(oRequest("ModifyDate").Item)
		Else
			aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION)) Then
		If Len(oRequest("PayrollDate").Item) > 0 Then
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = CInt(oRequest("PayrollDateYear").Item) & Right(("0" & oRequest("PayrollDateMonth").Item), Len("00")) & Right(("0" & oRequest("PayrollDateDay").Item), Len("00"))
		ElseIf Len(oRequest("PayrollDate").Item) > 0 Then
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = CLng(oRequest("PayrollDate").Item)
		Else
			aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_USER_ID_REVISION)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aPayrollRevisionComponent(N_USER_ID_REVISION) = CLng(oRequest("UserID").Item)
		Else
			aPayrollRevisionComponent(N_USER_ID_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(N_ACTIVE_REVISION)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aPayrollRevisionComponent(N_ACTIVE_REVISION) = CLng(oRequest("UserID").Item)
		Else
			aPayrollRevisionComponent(N_ACTIVE_REVISION) = -1
		End If
	End If

	If IsEmpty(aPayrollRevisionComponent(S_COMMENTS_REVISION)) Then
		If Len(oRequest("Comments").Item) > 0 Then
			aPayrollRevisionComponent(S_COMMENTS_REVISION) = oRequest("Comments").Item
		Else
			aPayrollRevisionComponent(S_COMMENTS_REVISION) = ""
		End If
	End If
	aPayrollRevisionComponent(S_COMMENTS_REVISION) = Left(aPayrollRevisionComponent(S_COMMENTS_REVISION), 2000)

	aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION) = True
	InitializePayrollRevisionComponent = Err.number
	Err.Clear
End Function

Function AddEmployeeAdjustmentForRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
'************************************************************
'Purpose: To add a adjustments and deductions for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeAdjustmentForRevision"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDate
	Dim iForPayrollIsActiveConstant
	
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para hacer la revisión."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "El número de empleado no existe."
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If aPayrollRevisionComponent(N_START_DATE_REVISION ) < lDate Then
				If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
					iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
				ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
					iForPayrollIsActiveConstant = N_PAYROLL_FOR_FEATURES
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
					iForPayrollIsActiveConstant = 0
				End If
				If VerifyPayrollIsActive(oADODBConnection, aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), iForPayrollIsActiveConstant, sErrorDescription) Then
					lErrorNumber = CheckExistencyOfEmployeeAdjustmentForRevision(aPayrollRevisionComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo agregar la información del reclamo de pago por ajustes y deducciones."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAdjustmentsLKP (EmployeeID, ConceptID, ConceptAmount, MissingDate, PaymentDate, ModifyDate, PayrollDate, BeneficiaryName, UserID, Active, AdjustmentType) Values (" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ", " & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ", " & aPayrollRevisionComponent(D_CONCEPT_AMOUNT_REVISION) & ", " & aPayrollRevisionComponent(N_START_DATE_REVISION) & ", 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) & ", '', " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, 1)", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					Else
						sErrorDescription = "El concepto ya fue registrado con anterioridad."
					End If
				Else
					lErrorNumber = -1
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "La fecha de nómina para la revisión de pago no puede ser mayor o igual a la fecha actual."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeAdjustmentForRevision = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To add a adjustments and deductions for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeRevision"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oItem
	Dim sErrorDescription1
	Dim iForPayrollIsActiveConstant

	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Or (aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o la quincena de revisión para insertar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "El número de empleado no existe."
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
			ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
				iForPayrollIsActiveConstant = N_PAYROLL_FOR_FEATURES
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				iForPayrollIsActiveConstant = 0
			End If
			If VerifyPayrollIsActive(oADODBConnection, aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), iForPayrollIsActiveConstant, sErrorDescription) Then
				For Each oItem In oRequest("PayrollRevision")
					If CLng(oItem) < aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) Then
						sErrorDescription = "No se pudo agregar la información de la revisión de nómina para el empleado."
						lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesRevisions(PayrollID, EmployeeID, StartPayrollID, UserID, AddDate, Comments) Values (" & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) & ", " & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ", " & oItem & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", '" & Replace(aPayrollRevisionComponent(S_COMMENTS_REVISION), "'", "´") & "')", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription)
					Else
						lErrorNumber = -1
						sErrorDescription = "La quincena " & GetDateFromSerialNumber(CLng(oItem)) & " es mayor o igual a la quincena de aplicación " & GetDateFromSerialNumber(aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION))
					End If
					If lErrorNumber <> 0 Then
						sErrorDescription1 = sErrorDescription1 & " " & sErrorDescription & ","
					End If
				Next
				If Len(sErrorDescription1) > 0 Then
					lErrorNumber = -1
					sErrorDescription1 = Left(sErrorDescription1, (Len(sErrorDescription1) - Len(","))) & "."
					sErrorDescription = "Registros incorrectos:</BR>" & sErrorDescription1
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "La quincena de aplicación no está activa."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeRevision = lErrorNumber
	Err.Clear
End Function

Function GetPayrollRevisions(oRequest, oADODBConnection, aPayrollRevisionComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the absences for
'         the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollRevisions"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) <> -1 Then
		sCondition = "And (EmployeesRevisions.EmployeeID=" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ")"
	Else
		sCondition = "And (EmployeesRevisions.EmployeeID=0)"
	End If

	sCondition  = Trim(sCondition)
	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) <> 1 Then sCondition  = "And " & sCondition
	End If

	'sCondition = sCondition & " And (EmployeesRevisions.UserID=" & aLoginComponent(N_USER_ID_LOGIN) &")"
	If aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) <> -1 Then
 		sErrorDescription = "Seleccione un empleado para consultar sus registros de revisión."
	Else
		sErrorDescription = "No se pudo obtener la información de los registros de revisión de empleados."
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesRevisions.*, EmployeeName, EmployeeLastName, EmployeeLastName2, UserName, UserLastName From EmployeesRevisions, Employees, Users Where (EmployeesRevisions.EmployeeID=Employees.EmployeeID) And (EmployeesRevisions.UserID=Users.UserID)" & sCondition & " Order By PayrollID Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetPayrollRevisions = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an employee's child from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeRevision"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iForPayrollIsActiveConstant

	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Or (aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = -1) Or (aPayrollRevisionComponent(N_START_DATE_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o la quincena de revisión o la quincena en que aplica la revisión para eliminar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
			iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
		ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
			iForPayrollIsActiveConstant = N_PAYROLL_FOR_FEATURES
		ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
			iForPayrollIsActiveConstant = 0
		End If
		If VerifyPayrollIsActive(oADODBConnection, aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), iForPayrollIsActiveConstant, sErrorDescription) Then
			sErrorDescription = "No se pudo eliminar la el registro de revisión de pagos para el empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRevisions Where (EmployeeID=" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (PayrollID=" & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) & ") And (StartPayrollID=" & aPayrollRevisionComponent(N_START_DATE_REVISION) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)			
		Else
			lErrorNumber = -1
			sErrorDescription = "No se puede eliminar el registro de revisión debido a que la quincena de aplicación ya fue cerrada."
		End If
	End If

	RemoveEmployeeRevision = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeAdjustmentForRevision(aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a concept exists in Adjustments database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeAdjustmentForRevision"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aEmployeeComponent(N_ID_EMPLOYEE) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAdjustmentsLKP Where (EmployeeID=" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") And (MissingDate=" & aPayrollRevisionComponent(N_START_DATE_REVISION) & ") And (AdjustmentType=1)", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sErrorDescription = "El registro ya fue registrado con anterioridad."
				lErrorNumber = L_ERR_DUPLICATED_RECORD
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeAdjustmentForRevision = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeRevision(aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a payroll revision for a employee in an application date
'		  exists in EmployeesRevisions database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeRevision"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Or (aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) = -1) Or (aPayrollRevisionComponent(N_START_DATE_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o la quincena de revisión o la quincena en que aplica la revisión para verificar si ya existe."
	Else
		sErrorDescription = "No se pudo revisar la existencia de la revisión de pago para el empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRevisions Where (EmployeeID=" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (PayrollID=" & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) &") And (StartPayrollID=" & aPayrollRevisionComponent(N_START_DATE_REVISION) & ")", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sErrorDescription = "Ya exise registrada la revisión del empleado para la quincena " & aPayrollRevisionComponent(N_START_DATE_REVISION) & " para ser aplicado en " & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION)
				aPayrollRevisionComponent(B_IS_DUPLICATED_REVISION) = True
				lErrorNumber = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeRevision = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollRevisionForm(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, aPayrollRevisionComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollRevisionForm"
	Dim sNames
	Dim aRelatedAbsences
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAction

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var payrollCont=0;" & vbNewLine
			Response.Write "function CheckRevisionFields(oForm) {" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.PayrollRevision);" & vbNewLine
				Response.Write "if (GetSelectedItems(oForm.PayrollRevision) == ''){" & vbNewLine
					Response.Write "alert('Seleccione al menos una quincena para revisión.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine

				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckRevisionFields" & vbNewLine

			Response.Write "function ShowHideRevisionFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.PayrollRevisionFrm" & vbNewLine
				If Not B_ISSSTE Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (sValue == 0) {" & vbNewLine
							Response.Write "HideDisplay(document.all['PayrollsForRevisionDiv']);" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "ShowDisplay(document.all['PayrollsForRevisionDiv']);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
			Response.Write "} // End of ShowHideRevisionFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		If aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1 Then
			Response.Write "<FORM NAME=""PayrollRevisionFrm"" ID=""PayrollRevisionFrm"" ACTION=""UploadInfo.asp"" METHOD=""GET"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PayrollRevision"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR /><BR />"
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""PayrollRevision"" ID=""PayrollRevisionBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" />"
				End If
			Response.Write "</FORM>"
		Else
			lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", "PayrollRevision", 400, lReasonID, "Revisión de nóminas", "Registre las revisiones de pagos a un empleado diferente", sErrorDescription)
					Response.Write "<FORM NAME=""PayrollRevisionFrm"" ID=""PayrollRevisionFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckRevisionFields(this)"">"
						If Len(oRequest("RevisionChange").Item) > 0 Then
							lErrorNumber = GetRevision(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RevisionChange"" ID=""RevisionChangeHdn"" VALUE=""" & oRequest("RevisionChange").Item & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aLoginComponent(N_USER_ID_LOGIN) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifyDate"" ID=""ModifyDateHdn"" VALUE=""" & Left(GetSerialNumberForDate(""), Len("00000000")) & """ />"
							aPayrollRevisionComponent(N_ACTIVE_REVISION) = 0
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PayrollRevision"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""4"" />"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
								If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
									Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
								Else
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
								End If
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de empleado:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "PositionTypes", aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Positions", aEmployeeComponent(N_POSITION_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						If Len(oRequest("RevisionChange").Item) > 0 Then
							Response.Write "<TR><TD COLSPAN=""2""><BR /></TD></TR>"
							Response.Write "<TR><TD COLSPAN=""2"">"
							If Len(oRequest("CancelRevision").Item) > 0 Then
								Call DisplayErrorMessage("Cancelar Revision", "Seleccione la quincena de aplicación para la cancelación.")
							Else
								Call DisplayErrorMessage("Justificar Revision", "Seleccione la quincena de aplicación.")
							End If
						Else
							Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación de la revisión:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""PayrollDate"" ID=""PayrollDate"" SIZE=""1"" CLASS=""Lists"">"
									If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
									ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
									ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And ((PayrollTypeID=1) Or (PayrollTypeID=4))", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
									End If
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
							Response.Write "</TR>"
						End If
						If False Then
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto para revisión:&nbsp;</FONT></TD>"
								Response.Write "<TD><SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									If Len(oRequest("RevisionChange").Item) > 0 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ")", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
									Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID IN (" & BENEFIT_CONCEPTS_FOR_PAYROLL & "))", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
									End If
								Response.Write "</SELECT></TD>"
							Response.Write "</TR>"
						End If
						Response.Write "</TABLE><BR />"
						Response.Write "<B>Seleccione la nómina a partir de la que se hará la revisión.</B><BR /><BR />"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function AddPayrollsToList(sYear, oForm) {" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "for (var i=0; i<=12; i++) {" & vbNewLine
										Response.Write "if (oForm.PayrollNumber[i].checked)" & vbNewLine
												Response.Write "AddItemToList(sYear + oForm.PayrollNumber[i].value, sYear + oForm.PayrollNumber[i].value, null, oForm.PayrollRevision);" & vbNewLine
										Response.Write "if (oForm.PayrollNumber[i + 12].checked) {" & vbNewLine
											Response.Write "if (oForm.PayrollNumber[i + 12].value == '0') {" & vbNewLine
												Response.Write "if ((parseInt(sYear) % 4) == 0)" & vbNewLine
													Response.Write "AddItemToList(sYear + '0229', sYear + '0228', null, oForm.PayrollRevision);" & vbNewLine
												Response.Write "else" & vbNewLine
													Response.Write "AddItemToList(sYear + '0228', sYear + '0228', null, oForm.PayrollRevision);" & vbNewLine
											Response.Write "} else {" & vbNewLine
												Response.Write "if (oForm.PayrollNumber[i + 12].value != '0106')" & vbNewLine
													Response.Write "AddItemToList(sYear + oForm.PayrollNumber[i + 12].value, sYear + oForm.PayrollNumber[i + 12].value, null, oForm.PayrollRevision);" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "for (var i=0; i<=24; i++)" & vbNewLine
										Response.Write "oForm.PayrollNumber[i].checked = false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "} // End of AddPayrollsToList" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine

						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
								For iIndex = 1 To 23 Step 2
									Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""PayrollNumber"" ID=""PayrollNumber"" VALUE=""" & Right(("0" & (Int(iIndex / 2) + 1)), Len("00")) & "15"" /> " & iIndex & "<BR />"
								Next
								Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""PayrollNumber"" ID=""PayrollNumber"" VALUE=""0106"" />Nómina de reyes<BR />"
							Response.Write "</FONT></TD>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
								For iIndex = 2 To 24 Step 2
									Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""PayrollNumber"" ID=""PayrollNumber"" VALUE="""
										Select Case iIndex
											Case 2,6,10,14,16,20,24
												Response.Write Right(("0" & Int(iIndex / 2)), Len("00")) & "31"
											Case 4
												Response.Write "0"
											Case 8,12,18,22
												Response.Write Right(("0" & Int(iIndex / 2)), Len("00")) & "30"
										End Select
									Response.Write """ /> " & iIndex & "<BR />"
								Next
							Response.Write "</FONT></TD>"
							Response.Write "<TD VALIGN=""TOP"">&nbsp;&nbsp;<SELECT NAME=""PayrollYear"" ID=""PayrollYear"" SIZE=""1"" CLASS=""Lists"">"
								For iIndex = Year(Date()) To 2008 Step -1
									Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
								Next
							Response.Write "</SELECT>&nbsp;&nbsp;</TD>"
							Response.Write "<TD VALIGN=""TOP"">"
								Response.Write "<A HREF=""javascript: AddPayrollsToList(document.PayrollRevisionFrm.PayrollYear.value, document.PayrollRevisionFrm)""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar nóminas para revisión"" BORDER=""0"" HSPACE=""5"" /></A>"
								Response.Write "<BR /><BR />"
								Response.Write "<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.PayrollRevisionFrm.PayrollRevision)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A>"
							Response.Write "</TD>"
							Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""PayrollRevision"" ID=""PayrollRevisionLst"" SIZE=""12"" MULTIPLE=""1"" CLASS=""Lists""></SELECT></TD>"
						Response.Write "</TABLE><BR />"

						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR><TD ALIGN=""TOP"" VALIGN=""TOP"">"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Observaciones:</FONT></TD>"
								Response.Write "<TD ALIGN=""TOP"" VALIGN=""TOP"">"
								Response.Write "<TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aPayrollRevisionComponent(S_COMMENTS_REVISION) & "</TEXTAREA>"
							Response.Write "</TD></TR>"
						Response.Write "</TABLE><BR />"

						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "ShowHideRevisionFields(1);" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine

						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""340"" HEIGHT=""1"" /><BR /><BR />"
						If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
							If InStr(1, sAction, "Employees") = 0 Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar Revisión"" CLASS=""Buttons"" />"
							End If
						ElseIf Len(oRequest("Delete").Item) > 0 Then
							If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAbsenceWngDiv']); AbsencesFrm.Remove.focus()"" />"
						Else
							If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""ShowDisplay(document.all['RemoveAbsenceWngDiv']);"" />"
						End If
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						If InStr(1, sAction, "Employees") = 0 Then
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Absences'"" />"
						End If
						Response.Write "<BR /><BR />"
					Response.Write "</FORM>"
				End If
			End If
		End If
	End If

	DisplayPayrollRevisionForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollsForRevisionSelectionList(oRequest, oADODBConnection, iPayrollIsClosed, aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, aPayrollRevisionComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollsForRevisionSelectionList"
	Dim oRecordset
	Dim lErrorNumber
	Dim sAction
	Dim bComponentInitialized

	bComponentInitialized = aPayrollRevisionComponent(B_COMPONENT_INITIALIZED_REVISION)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
	End If

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para hacer la revisión."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (IsClosed=" & iPayrollIsClosed & ") Order by PayrollID Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""PayrollRevision"" ID=""" & CStr(oRecordset.Fields("PayrollID").Value) & """ VALUE=""" & CStr(oRecordset.Fields("PayrollID").Value) & """"
					Response.Write " />" & CStr(oRecordset.Fields("PayrollID").Value) & " - " & CStr(oRecordset.Fields("PayrollName").Value) & "<BR />"
					oRecordset.MoveNext
				Loop
				oRecordset.Close
			Else
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	DisplayPayrollsForRevisionSelectionList = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollRevisionTable(oRequest, oADODBConnection, bForExport, aPayrollRevisionComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollRevisionTable"
	Dim oRecordset
	Dim iRecordCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNames
	Dim lErrorNumber

	lErrorNumber = GetPayrollRevisions(oRequest, oADODBConnection, aPayrollRevisionComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			'If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("Acciones,No. Empleado,Nombre,Q.Aplicación,Q.Inicio de Revisión,Usuario que capturo,F.Registro,Observaciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("120,100,300,200,200,200,300,400", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,LEFT", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("No. Empleado,Nombre,Q.Aplicación,Q.Inicio de Revisión,Usuario que capturo,F.Registro,Observaciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,300,200,200,200,300,400", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,LEFT", ",", -1, vbBinaryCompare)
				End If
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split("CENTER,CENTER,,,,CENTER,,LEFT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					sBoldBegin = ""
					sBoldEnd = ""
					sRowContents = ""
					If (Not bForExport) And (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=PayrollRevision&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&PayrollDate=" & CStr(oRecordset.Fields("PayrollID").Value) & "&StartPayrollID=" & CStr(oRecordset.Fields("StartPayrollID").Value) & "&Remove=1"">"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar revisión"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR
					End If
					sRowContents = sRowContents & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd & sFontEnd
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartPayrollID").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("AddDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					If (Not IsNull(oRecordset.Fields("Comments").Value)) And (Len(oRecordset.Fields("Comments").Value) > 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value)) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Ninguna") & sBoldEnd & sFontEnd
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1 Then
				sErrorDescription = "Introduzca un número de empleado para consultar las nóminas para revisión que tiene registradas. Si requiere un reporte especifico generelo desde el módulo de Reportes."
			Else
				sErrorDescription = "No se ha registrado revisión de nómina para este empleado."
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayPayrollRevisionTable = lErrorNumber
	Err.Clear
End Function

Function VerifyConceptIsActiveInPeriod(oRequest, oADODBConnection, aPayrollRevisionComponent, lAmount, sErrorDescription)
'************************************************************
'Purpose: To verify if concept was active in a
'         specific date 
'Inputs:  oRequest, oADODBConnection, aPayrollRevisionComponent
'Outputs: lAmount, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyConceptIsActiveInPeriod"
	Dim oRecordset
	Dim lErrorNumber
	Dim sAction
	Dim sQuery

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para hacer la revisión."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		Select Case aPayrollRevisionComponent(N_CONCEPT_ID_REVISION)
			Case 5, 19, 22, 24, 26, 32, 45, 46, 50, 69, 63, 67, 70, 72, 73, 76, 77, 93, 94, 104, 146
				sQuery = "Select * From EmployeesConceptsLKP Where (EmployeeID =" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") And (EndDate>=" & aPayrollRevisionComponent(N_START_DATE_REVISION) & ") And (Active=1) Order by StartDate Desc"
			Case 9, 14
				sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID =" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") And (OcurredDate=" & aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) & ") And  Order by OcurredDate Desc"
			Case Else
				sQuery = "Select * From EmployeesConceptsLKP Where (EmployeeID =" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") And (StartDate=" & aPayrollRevisionComponent(N_MODIFY_DATE_REVISION) & ") Order by StartDate Desc"
		End Select
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Select Case aPayrollRevisionComponent(N_CONCEPT_ID_REVISION)
					Case 5, 19, 22, 24, 26, 32, 45, 46, 50, 69, 63, 67, 70, 72, 73, 76, 77, 93, 94, 104, 146
						lAmount = CSng(oRecordset.Fields("ConceptAmount").Value)
					Case 9, 14
						lAmount = CSng(oRecordset.Fields("AbsenceHours").Value)
					Case Else
						lAmount = CSng(oRecordset.Fields("ConceptAmount").Value)
				End Select
			End If
			VerifyConceptIsActiveInPeriod = (Not oRecordset.EOF)
		Else
			sErrorDescription = "Error al verificar si el empleado tiene registrado el concepto."
			VerifyConceptIsActiveInPeriod = False
		End If
	End If

	Err.Clear
End Function

Function VerifyConceptInPayrrollForEmployee(oRequest, oADODBConnection, aPayrollRevisionComponent, lPayrollYear, lAmount, sErrorDescription)
'************************************************************
'Purpose: To verify if payroll concept was paid
'         to an employee
'Inputs:  oRequest, oADODBConnection, aPayrollRevisionComponent
'Outputs: lAmount, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyConceptInPayrrollForEmployee"
	Dim sNames
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAction

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para hacer la revisión."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payroll_" & lPayrollYear & " Where (RecordDate=" & aPayrollRevisionComponent(N_START_DATE_REVISION) & ") And (EmployeeID =" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") Order by RecordDate Desc", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lAmount = 0
			Else
				lAmount = CSng(oRecordset.Fields("ConceptAmount").Value)
			End If
			VerifyConceptInPayrrollForEmployee = (Not oRecordset.EOF)
		Else
			sErrorDescription = "Error al verificar si el empleado tiene registrado el concepto."
			VerifyConceptInPayrrollForEmployee = False
		End If
	End If

	Err.Clear
End Function

Function PayrollRevision(oRequest, oADODBConnection, aPayrollRevisionComponent, lPayrollYear, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, aPayrollRevisionComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollRevisionForm"
	Dim sNames
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAction

	If (aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para hacer la revisión."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollRevisionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payroll_" & lPayrollYear & " Where (RecordDate=" & aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION) & ") And (EmployeeID =" & aPayrollRevisionComponent(N_EMPLOYEE_ID_REVISION) & ") And (ConceptID=" & aPayrollRevisionComponent(N_CONCEPT_ID_REVISION) & ") Order by PayrollID Desc", "PayrollRevisionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			VerifyExistenceOfEmployeesBeneficiary = (Not oRecordset.EOF)
		Else
			sErrorDescription = "Error al verificar si el empleado tiene registradas revisiones."
			VerifyExistenceOfEmployeesBeneficiary = False
		End If
	End If

	DisplayPayrollRevisionForm = lErrorNumber
	Err.Clear
End Function
%>