<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/BanamexCensusComponent.asp" -->
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/AbsenceComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/EmployeesLib.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/JobsLib.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<!-- #include file="Libraries/PaymentsLib.asp" -->
<!-- #include file="Libraries/PaymentComponent.asp" -->
<!-- #include file="Libraries/AlimonyTypeComponent.asp" -->
<!-- #include file="Libraries/PayrollRevisionComponent.asp" -->
<!-- #include file="Libraries/ProfessionalRiskComponent.asp" -->
<!-- #include file="Libraries/PayrollResumeForSarComponent.asp" -->
<%
Dim bAction
Dim sAction
Dim iStep
Dim sFileName
Dim oItem
Dim lEmployeeTypeID
Dim sNames
Dim sMessage
Dim lEmployeeID
Dim lReasonID
Dim lSuccess
Dim sThirdConcept
Dim sError
Dim lStartDate
Dim lEndDate
Dim sOriginalFileName
Dim sEmployeeName
Dim sEmployeeLastName
Dim sEmployeeLastName2
Dim sRFC
Dim sCURP
Dim sThirdFileName
Dim sAbsenceShortName
Dim oJobRecordset
Dim oPayrollRecordset
Dim alRecordID
Dim iIndex
Dim bUploadFile
Dim sUploadFile
Dim sUploadFileError

sError = ""
lReasonID = 1

If CLng(oRequest("SubSectionID").Item) > 0 Then Response.Cookies("SIAP_SubSectionID") = CInt(oRequest("SubSectionID").Item)
If Len(oRequest("AbsenceShortName").Item) > 0 Then sAbsenceShortName = CStr(oRequest("AbsenceShortName").Item)
If Len(oRequest("ReasonID").Item) > 0 Then lReasonID = CLng(oRequest("ReasonID").Item)
If Len(oRequest("EmployeeID").Item) > 0 Then lEmployeeID = CLng(oRequest("EmployeeID").Item)
If Len(oRequest("Success").Item) > 0 Then lSuccess = 1
If Len(oRequest("ErrorDescription").Item) > 0 Then sError = CStr(oRequest("ErrorDescription").Item)
If Len(oRequest("ThirdConcept").Item) > 0 Then 
	sThirdConcept = CStr(oRequest("ThirdConcept").Item)
	lReasonID = 300
End If
If Len(oRequest("OriginalFile").Item) > 0 Then sOriginalFileName = CStr(oRequest("OriginalFile").Item)
If Len(oRequest("EmployeeName").Item) > 0 Then sEmployeeName = CStr(oRequest("EmployeeName").Item)
If Len(oRequest("EmployeeLastName").Item) > 0 Then sEmployeeLastName = CStr(oRequest("EmployeeLastName").Item)
If Len(oRequest("EmployeeLastName2").Item) > 0 Then sEmployeeLastName2 = CStr(oRequest("EmployeeLastName2").Item)
If Len(oRequest("RFC").Item) > 0 Then sRFC = CStr(oRequest("RFC").Item)
If Len(oRequest("CURP").Item) > 0 Then sCURP = CStr(oRequest("CURP").Item)

If Len(oRequest("sEmployeeName").Item) > 0 Then sEmployeeName = CStr(oRequest("sEmployeeName").Item)
If Len(oRequest("sEmployeeLastName").Item) > 0 Then sEmployeeLastName = CStr(oRequest("sEmployeeLastName").Item)
If Len(oRequest("sEmployeeLastName2").Item) > 0 Then sEmployeeLastName2 = CStr(oRequest("sEmployeeLastName2").Item)
If Len(oRequest("sRFC").Item) > 0 Then sRFC = CStr(oRequest("sRFC").Item)
If Len(oRequest("sCURP").Item) > 0 Then sCURP = CStr(oRequest("sCURP").Item)
If Len(oRequest("UploadFile").Item) > 0 Then
	Select Case lReasonID
		Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
			sUploadFile = Server.MapPath(UPLOADED_PHYSICAL_PATH & "Prestaciones\" & CStr(oRequest("UploadFile").Item))
		Case 300
			sUploadFile = Server.MapPath(UPLOADED_PHYSICAL_PATH & "Discos\" & CStr(oRequest("UploadFile").Item))
			sOriginalFileName = CStr(oRequest("UploadFile").Item)
	End Select
	bUploadFile = True
End If
sAction = oRequest("Action").Item
If Len(oRequest("EmployeeTypeID").Item)>0 Then
	lEmployeeTypeID = CLng(oRequest("EmployeeTypeID").Item)
Else
	lEmployeeTypeID = -1
End If
iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)
Select Case lReasonID
	Case EMPLOYEES_EFFICIENCY_AWARD
		sFileName = SYSTEM_PHYSICAL_PATH & CStr(oRequest("FileName").Item)
	Case Else
		If StrComp(sAction,"ProcessForSar",vbBinaryCompare) = 0 Then
			sFileName = Server.MapPath(UPLOADED_PHYSICAL_PATH & oRequest("Load").Item & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt")
		Else
			sFileName = Server.MapPath(UPLOADED_PHYSICAL_PATH & sAction & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt")
		End If
End Select

If Len(oRequest("RawData").Item) > 0 Then
	lErrorNumber = SaveTextToFile(sFileName, oRequest("RawData").Item, sErrorDescription)
    
    If sAction = "ProcessForSar" Then
       FormatingTextTabColumns oRequest("Load").Item, oRequest("Load").Item & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt"
    End If
    
	If lErrorNumber = 0 Then
		Select Case sAction
			Case "ConceptsValues"
				Response.Redirect "UploadInfo.asp?Action=" & sAction & "&EmployeeTypeID=" & lEmployeeTypeID & "&Step=" & iStep
			Case "ProcessForSar"
				Response.Redirect "UploadInfo.asp?Action=" & sAction & "&ReasonID=" & lReasonID & "&Step=" & iStep & "&Load=" & oRequest("Load").Item
			Case Else
				Select Case lReasonID
					Case 300
						Response.Redirect "UploadInfo.asp?Action=" & sAction & "&ReasonID=" & lReasonID & "&Step=" & iStep & "&ThirdConcept=" & sThirdConcept
					Case Else
						Response.Redirect "UploadInfo.asp?Action=" & sAction & "&ReasonID=" & lReasonID & "&Step=" & iStep
				End Select
		End Select
	End If
End If

Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
Call InitializePaymentComponent(oRequest, aPaymentComponent)
Call InitializePayrollRevisionComponent(oRequest, aPayrollRevisionComponent)
Call InitializeConceptComponent(oRequest, aConceptComponent)

If Len(oRequest("IsBatch").Item) > 0 Then
	Select Case CInt(oRequest("ReasonID").Item)
		Case 14: lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
	End Select
Else
	If Len(oRequest("SaveConceptValue").Item) > 0 Then
		If lErrorNumber = 0 Then
			lErrorNumber = AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El puesto no existe en la base de datos."
		End If
		If lErrorNumber = 0 Then
			Response.Redirect "UploadInfo.asp?Action=ConceptsValues&Success=1&EmployeeTypeID=" & lEmployeeTypeID
		Else
			Response.Redirect "UploadInfo.asp?Action=ConceptsValues&Success=0&EmployeeTypeID=" & lEmployeeTypeID & "&ErrorDescription=" & sErrorDescription
		End If
	End If
	If Len(oRequest("SaveEmployeeConcept").Item) > 0 Then
		aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
		'aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
		lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			Response.Redirect "UploadInfo.asp?Action=" & sAction & "&Success=1"
		Else
			Response.Redirect "UploadInfo.asp?Action=" & sAction & "&Success=0&ErrorDescription=" & sErrorDescription
		End If
	End If
	If Len(oRequest("SaveEmployeeChildren").Item) > 0 And (StrComp(sAction, "ChildrenSchoolarships", vbBinaryCompare) <> 0) Then
		lErrorNumber = SaveEmployeeChildren(aEmployeeComponent, sErrorDescription)
	End If
	If Len(oRequest("ConceptValuesAction").Item) > 0 Then
		If (Len(oRequest("AuthorizationFile").Item) > 0) Then
			lErrorNumber = AddConceptsValuesFile(oRequest, oADODBConnection,  oRequest("sQuery").Item, aConceptComponent, sErrorDescription)
			sError = sErrorDescription
			If lErrorNumber = 0 Then
				sError = sError & "El tabulador de pago se registró exitosamente<BR />"
			Else
				sError = sError & "Error al registrar el tabulador de pago<BR />"
			End If
			Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=1&ErrorDescription=" & sError
		ElseIf (Len(oRequest("RemoveFile").Item) > 0) Then
			lErrorNumber = RemoveConceptsValuesFile(oRequest, oADODBConnection, oRequest("sQuery").Item, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("Remove").Item) > 0) Then
			For Each oItem In oRequest("RecordID")
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oItem)
				lErrorNumber = RemoveConceptValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					sMessage = sMessage & ", " & sErrorDescription
				End If
			Next
			If Len(sMessage) = 0 Then
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=0&ErrorDescription=" & sMessage
			End If
		ElseIf (Len(oRequest("Apply").Item) > 0) Then
			alRecordID = oRequest("RecordID")
			alRecordID = Split(alRecordID, ",")
			For iIndex = 0 To UBound(alRecordID)
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(alRecordID(iIndex))
				lErrorNumber = SetActiveForConceptsValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					sMessage = sMessage & ", " & sErrorDescription
				End If
			Next
			If Len(sMessage) = 0 Then
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("ChangeEndDateButton").Item) > 0) Then
			alRecordID = oRequest("RecordID")
			alRecordID = Split(alRecordID, ",")
			For iIndex = 0 To UBound(alRecordID)
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(alRecordID(iIndex))
				lErrorNumber = ModifyEndDateForConceptsValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					sMessage = sMessage & ", " & sErrorDescription
				End If
			Next
			If Len(sMessage) = 0 Then
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=" & lEmployeeTypeID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		End If
	End If
	If Len(oRequest("SaveEmployeesMovements").Item) > 0 Then
		If (Len(oRequest("ActiveConcept").Item) > 0) Then
			Select Case lReasonID
				Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
					lErrorNumber = SetActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_GLASSES, EMPLOYEES_PROFESSIONAL_DEGREE
					lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case CANCEL_EMPLOYEES_CONCEPTS
					lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case Else
			End Select
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("DeActiveConcept").Item) > 0) Then
			Select Case lReasonID
				Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
					lErrorNumber = SetDeActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_GLASSES, EMPLOYEES_PROFESSIONAL_DEGREE
					lErrorNumber = SetDeActiveForEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case Else
			End Select
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("MoveBeneficiaryUp").Item) > 0) Then
			lErrorNumber = MoveEmployeeBeneficiaryUp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("RemoveMotion").Item) > 0) Then
			lErrorNumber = RemoveEmployeeForValidation(oRequest, oADODBConnection, "EmployeesMovements", aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = RemoveEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&MovementsSuccess=1&ErrorDescription=El movimiento fue cancelado exitosamente."
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&MovementsSuccess=0&ErrorDescription=Ocurrió un error al cancelar el movimiento."
			End If
		ElseIf (Len(oRequest("CancelMotion").Item) > 0) Then
			Select Case lReasonID
				Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
					lErrorNumber = CancelEmployeeConcept(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo cancelar la información."
					End If
				Case -89, EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, EMPLOYEES_FOR_RISK, EMPLOYEES_ANTIQUITIES, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CONCEPT_08, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_BENEFICIARIES_DEBIT, EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
					lErrorNumber = RemoveEmployeeForValidation(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
					End If
				Case EMPLOYEES_THIRD_CONCEPT
					lErrorNumber = RemoveEmployeeCredit(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo eliminar la información."
					End If
				Case -58
					lErrorNumber = RemoveEmployeeForValidation(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
					End If
				Case EMPLOYEES_ADD_BENEFICIARIES
					lErrorNumber = RemoveEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo eliminar la información."
					End If
				Case EMPLOYEES_BANK_ACCOUNTS
					lErrorNumber = RemoveEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo eliminar la información."
					End If
				Case EMPLOYEES_CREDITORS
					lErrorNumber = RemoveEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo eliminar la información."
					End If
				Case EMPLOYEES_GRADE
					lErrorNumber = RemoveEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo eliminar la información."
					End If
				Case Else
					lErrorNumber = AddEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&MovementsSuccess=1&ErrorDescription=" & "El rechazo al movimiento fue registrado exitosamente."
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&MovementsSuccess=1&ErrorDescription=" & "Ocurrió un error al registrar el rechazo al movimiento del empleado."
					End If
			End Select
		ElseIf (Len(oRequest("Modify").Item) > 0) Then
			Select Case aEmployeeComponent(N_REASON_ID_EMPLOYEE)
				Case 57, 58
					If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
						aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
						lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						End If
						aEmployeeComponent(N_ZONE_ID_EMPLOYEE) = aJobComponent(N_ZONE_ID_JOB)
						aEmployeeComponent(N_POSITION_ID_EMPLOYEE) = aJobComponent(N_POSITION_ID_JOB)
						aEmployeeComponent(N_AREA_ID_EMPLOYEE) = aJobComponent(N_AREA_ID_JOB)
					End If
					aEmployeeComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE) = False
					aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = False
					lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	'				If lReasonID = 57 Then
	'					lErrorNumber = ModifyEmployeeHistoryList(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	'				End If
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
					End If
				Case Else
					Select Case lReasonID
						Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
							lErrorNumber = ModifyEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						Case -89, EMPLOYEES_ANTIQUITIES, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_CONCEPT_7S, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
							lErrorNumber = ModifyEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						Case CANCEL_EMPLOYEES_SSI
							lErrorNumber = CloseEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						Case Else
					End Select
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					End If
			End Select
        ElseIf (Len(oRequest("ModifyChildren").Item) > 0)  Then
            lErrorNumber = ModifyChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
            If lErrorNumber = 0 Then
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
			End If
        ElseIf (Len(oRequest("Remove").Item) > 0)  Then
            lErrorNumber = DeleteChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
            If lErrorNumber = 0 Then
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
			End If
        ElseIf (Len(oRequest("AddChild").Item) > 0) Then
            lErrorNumber = AddChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
            If lErrorNumber = 0 Then
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
			    Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo modificar la información."
			End If
		ElseIf (Len(oRequest("Add").Item) > 0) Then
			Dim iEmployeeIDTemp
			Select Case lReasonID
				'Case ALIMONY_TYPES
				'	lErrorNumber = AddAlimonyTypes(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_GRADE
					lErrorNumber = AddEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
					lErrorNumber = AddDocumentsForLicenses(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
					aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,3"
					aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
					'lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case -89, EMPLOYEES_FOR_RISK, EMPLOYEES_ANTIQUITIES, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CONCEPT_08, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_MONTHAWARD, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_CONCEPT_7S, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
					Select Case lReasonID
						Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT
							aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
					End Select
					aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
					lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_NIGHTSHIFTS
					Dim sErrorDescription1
					Dim sNightShiftDates
					Dim iNightShifts
					iNightShifts = 0
					If Len(oRequest("OcurredDates").Item) > 0 Then
						For Each oItem In oRequest("OcurredDates")
							iNightShifts = iNightShifts + 1
							sNightShiftDates = sNightShiftDates & CStr(oItem) & ","
						Next
						If InStr(1, Right(sNightShiftDates, Len(","), ",")) Then
							sNightShiftDates = Left(sNightShiftDates, (Len(sNightShiftDates) - Len(",")))
						End If
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(iNightShifts / 6)
						aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = sNightShiftDates
						aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
						aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
						lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
				Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
					lErrorNumber = CloseEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_SUNDAYS
					If CInt(oRequest("SundayChange").Item) Then
						lErrorNumber = ModifyEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						If Len(oRequest("OcurredDates").Item) > 0 Then
							For Each oItem In oRequest("OcurredDates")
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oItem)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
								lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									sMessage = sMessage & sErrorDescription
								End If
							Next
							If Len(sMessage) > 0 Then
								lErrorNumber = -1
								sErrorDescription = sMessage
							End If
						Else
							lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					End If
				Case EMPLOYEES_EXTRAHOURS
					If CInt(oRequest("SundayChange").Item) Then
						lErrorNumber = ModifyEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						If Len(oRequest("OcurredDates").Item) > 0 Then
							For Each oItem In oRequest("OcurredDates")
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oItem)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
								lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									sMessage = sMessage & sErrorDescription
								End If
							Next
							If Len(sMessage) > 0 Then
								lErrorNumber = -1
								sErrorDescription = sMessage
							End If
						Else
							lErrorNumber = AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					End If
				Case EMPLOYEES_MOTHERAWARD
					aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
					aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
					lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case EMPLOYEES_ADD_BENEFICIARIES
					Dim iModifyBeneficiary
					iModifyBeneficiary = CInt(oRequest("BeneficiaryChange").Item)
					iEmployeeIDTemp = aEmployeeComponent(N_ID_EMPLOYEE)
					If iModifyBeneficiary Then
						lErrorNumber = ModifyEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						lErrorNumber = AddEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
					aEmployeeComponent(N_ID_EMPLOYEE) = iEmployeeIDTemp
				Case EMPLOYEES_CREDITORS
					iModifyBeneficiary = CInt(oRequest("BeneficiaryChange").Item)
					iEmployeeIDTemp = aEmployeeComponent(N_ID_EMPLOYEE)
					If iModifyBeneficiary Then
						lErrorNumber = ModifyEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						lErrorNumber = AddEmployeeCreditors(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
					aEmployeeComponent(N_ID_EMPLOYEE) = iEmployeeIDTemp
				Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_BENEFICIARIES_DEBIT
					Dim iModifyCredit
					iModifyCredit = CInt(oRequest("CreditChange").Item)
					If iModifyCredit Then
						lErrorNumber = ModifyEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) = CDbl(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) * aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE))
						aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
						If lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT Then
							aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = 86
							aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 0
						End If
						If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0 Then
							If aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) > 0 Then 
								lErrorNumber = GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							Else
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
							End If
						End If
						lErrorNumber = AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
				Case EMPLOYEES_BANK_ACCOUNTS
					Dim iModifyBankAccount
					Dim iChequeAccount
					iModifyBankAccount = CInt(oRequest("BankAccountChange").Item)
					iChequeAccount = CInt(oRequest("Cheque").Item)
					If iChequeAccount Then
						aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = "."
					End If
					If iModifyBankAccount Then
						lErrorNumber = ModifyEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					Else
						aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
						lErrorNumber = AddEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					End If
				Case EMPLOYEES_EXCENT
			End Select
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("Authorization").Item) > 0) Then
			Select Case aEmployeeComponent(N_REASON_ID_EMPLOYEE)
				Case -58
					lErrorNumber = SetActiveForEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=No se pudo autorzar el reclamo"
					End If
				Case EMPLOYEES_FOR_RISK, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_CONCEPT_08
					lErrorNumber = ModifyEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					sError = sErrorDescription
				Case -89, EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, EMPLOYEES_ANTIQUITIES, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_CONCEPT_7S, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
					aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
					'lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					'lErrorNumber = AddEmployeeConceptForValidation(oRequest, oADODBConnection, "EmployeesSafeSeparation", aEmployeeComponent, sErrorDescription)
					lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					End If
				Case EMPLOYEES_BENEFICIARIES_DEBIT
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Credits Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")" , "UploadInfo.asp", "Active_EMPLOYEES_BENEFICIARIES_DEBIT", 0, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0"
					End If
				Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
					'lErrorNumber = AddEmployeeAbsencesForValidation(oRequest, oADODBConnection, "EmployeesSafeSeparation", aEmployeeComponent, sErrorDescription)
					lErrorNumber = SetActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=0&ErrorDescription=" & sErrorDescription
					End If
				Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_THIRD_PROCESS, EMPLOYEES_BENEFICIARIES_DEBIT
					'lErrorNumber = SetActiveForEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					lErrorNumber = SetActiveForEmployeeCreditFile(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
						Else
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1"
						End If
					Else
						If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
						Else
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0&ErrorDescription=" & sErrorDescription
						End If
					End If
				Case EMPLOYEES_BANK_ACCOUNTS
					lErrorNumber = SetActiveForEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					End If
				Case EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_CREDITORS
					lErrorNumber = SetActiveForEmployeeBeneficiary(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					End If
				Case EMPLOYEES_GRADE
					lErrorNumber = SetActiveForEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					End If
			End Select
		End If
		If Len(oRequest("SaveEmployeesAdjustments").Item) Then
			lErrorNumber = AddEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			sError = sErrorDescription
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58&Success=1&EmployeeID=" & lEmployeeID
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58&Success=0&EmployeeID=" & lEmployeeID & "&ErrorDescription=" & sError
			End If
		ElseIf Len(oRequest("AuthorizationFile").Item) > 0 Then
			If lReasonID >= 0 Then
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
				lErrorNumber = AddEmployeeMovementFile(oRequest, oADODBConnection, oRequest("sQuery").Item, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
				sError = sErrorDescription
				If lErrorNumber = 0 Then
					sError = sError & "El movimiento del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " se registró exitosamente<BR />"
				Else
					sError = sError & "Error al registrar el movimiento del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " <BR />"
				End If
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&MovementsSuccess=1&ErrorDescription=" & sError
			Else
				Select Case	lReasonID
					Case EMPLOYEES_THIRD_PROCESS
						sThirdFileName = CStr(oRequest("ConceptFileName").Item)
						lErrorNumber = AddEmployeeConceptsFile(oRequest, oADODBConnection, oRequest("sQuery").Item, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
						sError = sErrorDescription
						If lErrorNumber = 0 Then
							sError = sError & "Los conceptos se registraron exitosamente<BR />"
							Call RemoveUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, sThirdFileName, sErrorDescription)
						Else
							sError = sError & "Error al registrar el concepto del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " <BR />"
							Call RemoveUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, sThirdFileName, sErrorDescription)
						End If
						Response.Redirect "UploadInfo.asp?Action=ThirdUploadMovements&ReasonID=" & lReasonID & "&MovementsSuccess=1&ErrorDescription=" & sError
					Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
						lErrorNumber = SetActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						sError = sErrorDescription
						If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&ErrorDescription=" & sError
						Else
							Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sError
						End If
					Case Else
						lErrorNumber = AddEmployeeConceptsFile(oRequest, oADODBConnection, oRequest("sQuery").Item, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
						sError = sErrorDescription
						Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sError
				End Select
			End If
		ElseIf Len(oRequest("RemoveFile").Item) > 0 Then
			Select Case lReasonID
				Case -58
					lErrorNumber = RemoveEmployeeConceptsFile(oRequest, oADODBConnection, oRequest("sQuery").Item, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
					sError = sErrorDescription
					If lErrorNumber = 0 Then
						sError = sError & "Los reclamos de pago en proceso de validación seleccionados se eliminaron exitosamente<BR />"
					Else
						sError = sError & "Error al eliminar los reclamos de pago en proceso de validación<BR />"
					End If
					Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&Success=1&ErrorDescription=" & sError
				Case EMPLOYEES_THIRD_PROCESS
					sThirdFileName = CStr(oRequest("ConceptFileName").Item)
					lErrorNumber = RemoveEmployeeConceptsFile(oRequest, oADODBConnection, oRequest("sQuery").Item, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
					sError = sErrorDescription
					If lErrorNumber = 0 Then
						sError = sError & "Los conceptos fueron eliminados exitosamente<BR />"
						Call RemoveUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, sThirdFileName, sErrorDescription)
					Else
						sError = sError & "Error al eliminar los conceptos del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " <BR />"
						Call RemoveUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, sThirdFileName, sErrorDescription)
					End If
					Response.Redirect "UploadInfo.asp?Action=ThirdUploadMovements&ReasonID=" & lReasonID & "&MovementsSuccess=1&ErrorDescription=" & sError
			End Select
		Else
			If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -1) And (lReasonID <> 1) And (lReasonID <> 14) And (lReasonID <> 26) And (lReasonID <> 53) And (lReasonID <> -64) And (lReasonID <> EMPLOYEES_CONCEPT_08) And (lReasonID <> EMPLOYEES_HONORARIUM_CONCEPT) Then
				If lReasonID = 28 Then
					lStartDate = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
					lEndDate = 30000000
				End If
				lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lReasonID = 28 Then
					aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lStartDate
					aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = lEndDate
					If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -3) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobId From JobsHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", "UploadInfo.asp", 0, sErrorDescription, oJobRecordset)
						aEmployeeComponent(N_JOB_ID_EMPLOYEE) = oJobRecordset.Fields("JobID").Value
					End If
					aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
				End If
				If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
					aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
					lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
					End If
				End If
			End If
				If lErrorNumber = 0 Then
					lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
					sError = sErrorDescription
					If lErrorNumber = 0 Then
						If (InStr(1, ",12,13,14,17,18,21,26,28,37,38,39,40,41,43,44,45,46,47,48,50,51,68,", "," & oRequest("ReasonID").Item & ",",vbBinaryCompare) > 0) Then
							If (Len(oRequest("Authorization").Item) > 0) Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID from Payrolls where PayrollTypeID = 1 and PayrollID > " & oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item, "EmployeeDisplayFormsComponent.asp", "UploadInfo.asp", 0, sErrorDescription, oPayrollRecordset)
								If Not oPayrollRecordset.EOF Then
									Do While Not oPayrollRecordset.EOF
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRevisions(PayrollID, EmployeeID, StartPayrollID, UserID, AddDate, Comments) Values(" & oRequest("EmployeePayrollDate").Item & ", " & oRequest("EmployeeID").Item & ", " & oPayrollRecordset.Fields("PayrollID").Value & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 'Registro agregado por el sistema a partir de la aplicación de un movimiento en el histórico del empleado.')", "EmployeeDisplayFormsComponent.asp", "UploadInfo.asp", 0, sErrorDescription, Null)
										oPayrollRecordset.MoveNext
									Loop
								End If
							End If
						End If
						If (InStr(1, ",1,2,3,4,5,6,8,10,21,26,50,51,62,63,66,101,102,103,104,105,106,", "," & oRequest("ReasonID").Item & ",",vbBinaryCompare) > 0) Then
							If (Len(oRequest("Authorization").Item) > 0) Then
								lErrorNumber = CheckForFractionatedPeriod(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
							End If
						End If
					End If
				End If
			If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 26) Then
				lEmployeeID = aEmployeeComponent(N_ID_EMPLOYEE)
			End If
			If lErrorNumber = 0 Then
				If (Len(oRequest("Register").Item) > 0) Then 
					sError = "El movimiento se registró exitosamente"
				End If
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&EmployeeID=" & lEmployeeID & "&MovementsSuccess=1&ErrorDescription=" & sError
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & "&EmployeeID=" & lEmployeeID & "&MovementsSuccess=0&ErrorDescription=" & sError
			End If
		End If
	End If
	If Len(oRequest("SaveAlimonyTypesMovements").Item) > 0 Then
		If (Len(oRequest("Add").Item) > 0) Then
			aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1
			Select Case lReasonID
				Case ALIMONY_TYPES
					lErrorNumber = AddAlimonyTypes(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				Case CREDITORS_TYPES
					lErrorNumber = AddCreditorTypes(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			End Select
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("Remove").Item) > 0) Then
			lErrorNumber = RemoveAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		ElseIf (Len(oRequest("Modify").Item) > 0) Then
			lErrorNumber = ModifyAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=1"
			Else
				Response.Redirect "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & lReasonID & "&Success=0&ErrorDescription=" & sErrorDescription
			End If
		End If
	End If
	If (Len(oRequest("SaveEmployeesAssignNumber").Item) > 0) Then
		lEmployeeID = aEmployeeComponent(N_ID_EMPLOYEE)
		lErrorNumber = CheckExistencyOfEmployeeRFC(aEmployeeComponent, sErrorDescription)
		If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
			aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID
			aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
			aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 0
			lErrorNumber = AddEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
					Response.Redirect "UploadInfo.asp?Action=EmployeesAssignTemporalNumber&Success=1&ReasonID=" & lReasonID
				Else
					Response.Redirect "UploadInfo.asp?Action=EmployeesAssignNumber&Success=1&ReasonID=" & lReasonID
				End If
			Else
				If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
					Response.Redirect "UploadInfo.asp?Action=EmployeesAssignTemporalNumber&Success=0&ReasonID=" & lReasonID & "&ErrorDescription=" & sErrorDescription
				Else
					Response.Redirect "UploadInfo.asp?Action=EmployeesAssignNumber&Success=0&ReasonID=" & lReasonID & "&ErrorDescription=" & sErrorDescription
				End If
			End If
		Else
			sErrorDescription = "El empleado con número " & aEmployeeComponent(N_ID_EMPLOYEE) & " tiene el mismo RFC, CURP o el mismo nombre completo."
			If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesAssignTemporalNumber&Success=0&ErrorDescription=" & sErrorDescription
			Else
				Response.Redirect "UploadInfo.asp?Action=EmployeesAssignNumber&Success=0&ErrorDescription=" & sErrorDescription & "&sEmployeeName=" & sEmployeeName & "&sEmployeeLastName=" & sEmployeeLastName & "&sEmployeeLastName2=" & sEmployeeLastName2 & "&sCURP=" & sCURP & "&sRFC=" & sRFC
			End If
		End If
	End If

	Select Case sAction
		Case "Jobs"		
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
				If lErrorNumber = 0 Then
					sErrorDescription = "La plaza número " & aJobComponent(N_ID_JOB) & " ha sido agregada con éxito"
					Response.Redirect "UploadInfo.asp?Action=Jobs&ReasonID=59&Success=1&ErrorDescription=" & sErrorDescription
				Else
					sErrorDescription = "No se pudo agregar la plaza"
					Response.Redirect "UploadInfo.asp?Action=Jobs&ReasonID=59&Success=0&ErrorDescription=" & sErrorDescription
				End If
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					aEmployeeComponent(N_ID_EMPLOYEE) = -1
					lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeeIDsFromJobs", aJobComponent(N_ID_JOB), "", "", aEmployeeComponent(N_ID_EMPLOYEE), sErrorDescription)
					If (lErrorNumber = 0) And (aEmployeeComponent(N_ID_EMPLOYEE) > -1) Then
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = aJobComponent(N_SERVICE_ID_JOB)
							aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) = aJobComponent(N_LEVEL_ID_JOB)
							lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					End If
				End If
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				If aJobComponent(N_ID_JOB) > -1 Then
					lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					If lErrorNumber = 0 Then lErrorNumber = RemoveJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				End If
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Else
				Call InitializeJobComponent(oRequest, aJobComponent)
			End If
		Case "ProfessionalRisk"
			If Len(oRequest("Add").Item) > 0 Then
			Else
				Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
			End If
	End Select

	Call GetEmployeesURLValues(oRequest, 1, bAction, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))
	If bAction Then
		lErrorNumber = DoEmployeesAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
		If lErrorNumber = 0 Then
			Select Case sAction
				Case "Absences"
					If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
						Response.Redirect "UploadInfo.asp?Action=Absences&Success=1&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&ActionConfirmation=1&AbsenceShortName=" & aAbsenceComponent(S_ABSENCE_SHORT_NAME_ABSENCE)
					Else
						Response.Redirect "UploadInfo.asp?Action=Absences&Success=1&ActionConfirmation=1&AbsenceShortName=" & aAbsenceComponent(S_ABSENCE_SHORT_NAME_ABSENCE)
					End If
				Case "PayrollRevision"
					If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
						Response.Redirect "UploadInfo.asp?Action=PayrollRevision&Success=1&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					Else
						Response.Redirect "UploadInfo.asp?Action=PayrollRevision&Success=1"
					End If
				Case "ApplyAbsences"
					Response.Redirect "UploadInfo.asp?Action=ApplyAbsences&Success=1"
				Case Else
			End Select
		Else
			Select Case sAction
				Case "Absences"
					If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
						Response.Redirect "UploadInfo.asp?Action=Absences&Success=0&ActionConfirmation=1&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&ErrorDescription=" & sErrorDescription & "&AbsenceShortName=" & aAbsenceComponent(S_ABSENCE_SHORT_NAME_ABSENCE)
					Else
						Response.Redirect "UploadInfo.asp?Action=Absences&Success=0&ErrorDescription=" & sErrorDescription & "&AbsenceShortName=" & aAbsenceComponent(S_ABSENCE_SHORT_NAME_ABSENCE)
					End If
				Case "PayrollRevision"
					If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
						Response.Redirect "UploadInfo.asp?Action=PayrollRevision&Success=0&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
					Else
						Response.Redirect "UploadInfo.asp?Action=PayrollRevision&Success=0&ErrorDescription=" & sErrorDescription
					End If
				Case "ApplyAbsences"
					Response.Redirect "UploadInfo.asp?Action=ApplyAbsences&Success=0&ErrorDescription=" & sErrorDescription
				Case Else
			End Select
		End If
		'sError = sErrorDescription
		'bError = (lErrorNumber <> 0)
		'If (lErrorNumber = 0) And (Len(oRequest("Remove").Item) > 0) Then
		'	bAction = False
		'End If
	End If

	Select Case sAction
		Case "Absences"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Incidencias"
			If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
			Else
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			End If
		Case "AlimonyTypes"
			Select Case lReasonID
				Case CREDITORS_TYPES
					aHeaderComponent(S_TITLE_NAME_HEADER) = "Tipos de descuento para pagar a acreedores"
				Case Else
					aHeaderComponent(S_TITLE_NAME_HEADER) = "Tipos de pensión alimenticia"
			End Select
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "ApplyAbsences"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Aplicación de Incidencias"
			If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Else
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			End If
		Case "ConceptsValues"
			If lEmployeeTypeID >= 0 Then
				Call GetNameFromTable(oADODBConnection, "EmployeeTypes", lEmployeeTypeID, "", "", sNames, "")
				aHeaderComponent(S_TITLE_NAME_HEADER) = sNames
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Resultado de la carga"
			End If
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Case "CreditFOVISSSTE"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "62. Crédito FOVISSSTE"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesAbsences"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Incidencias"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Case "EmployeesAccounts"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cuentas bancarias"
		Case "EmployeesAdditionalCompensation"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "08. Remuneración adicional (confianza)"
		Case "EmployeesAssignNumber"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Asignación de número de empleado"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "EmployeesAssignTemporalNumber"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Asignación de número temporal de empleado"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Case "EmployeesChanges"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Movimientos a los empleados"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "EmployeesCarLoan"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "73. Préstamo automóvil servidores públicos de mando superior"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesChildren"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Hijos de empleados"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesConcepts"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de conceptos de empleado"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesDrop"
			Call GetNameFromTable(oADODBConnection, "Reasons", lReasonID, "", "", sNames, "")
			aHeaderComponent(S_TITLE_NAME_HEADER) = sNames
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "EmployeesExtraHours"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "09. Remuneración por horas extraordinarias"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesFamilyDeath"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "42. Ayuda por muerte de familiar en primer grado"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesFONAC"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "77. FONAC"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesForRisk"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "04. Compensación por riesgos profesionales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesGlasses"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "20. Ayuda de anteojos"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesInactive"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Bloqueo de pagos"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesLicenses"
			Call GetNameFromTable(oADODBConnection, "Reasons", lReasonID, "", "", sNames, "")
			aHeaderComponent(S_TITLE_NAME_HEADER) = sNames
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "EmployeeMonthAward"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "49. Premio trabajador del mes"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesMovements"
			Call GetNameFromTable(oADODBConnection, "Reasons", lReasonID, "", "", sNames, "")
			If lReasonID = -91 Then ' No existe en la tabla Reasons este Id
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Aplicación de registros cargados por cada archivo"
			ElseIf lReasonID = EMPLOYEES_GRADE Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Calificación de empleados"
			Else
				Select Case CInt(Request.Cookies("SIAP_SubSectionID"))
					Case 2
						aHeaderComponent(S_TITLE_NAME_HEADER) = "Baja de los conceptos C9, 71 ó 72"
					Case Else
						aHeaderComponent(S_TITLE_NAME_HEADER) = sNames
				End Select
			End If
			Select Case CInt(Request.Cookies("SIAP_SectionID"))
				Case 1
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
				Case 2
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
				Case 4
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
				Case 7
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
				Case Else
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			End Select
		Case "EmployeesNew"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "101 Nuevo ingreso"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			lErrorNumber = GetEmployeeByStatus(oRequest, oADODBConnection, "1", "EmployeesNew", aEmployeeComponent, sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Redirect "UploadInfo.asp?Action=EmployeesAssignNumber"
			End If
		Case "EmployeesNightShifts"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "C2. Jornada nocturna adicional por día festivo (acumulada)"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesProfessionalDegree"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "43. Ayuda impresión de tesis"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesResumptions"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Reanudación de labores"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		'Case "EmployeesSafeSeparation", "EmployeesAddSafeSeparation"
		'	aHeaderComponent(S_TITLE_NAME_HEADER) = "SI. Seguro de separación y AE. Seguro adicional de separación individualizado"
		'	aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesSAR"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "79. SAR"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeeSports"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "67. Cuota deportivo"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "EmployeesSundays"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "14. Primas dominicales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "FONAC"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Entrada del archivo de FOVISSSTE"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		Case "Internship"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "14. Primas dominicales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "Jobs"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Agregar una nueva plaza"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			If lReasonID = 60 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Cambio de datos a las plazas"
			ElseIf lReasonID = 61 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Cambio de puesto a las plazas"
			End If
			Select Case CInt(Request.Cookies("SIAP_SectionID"))
		        Case 1
			        aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
                Case 3
                    aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		        Case 4
			        aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		        Case 7
			        aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		        Case Else
			        aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
            End Select
		Case "JobServices"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cambio masivo del servicio"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "MedicalAreas"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Carga del archivo que contiene la información UNIMED"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Case "MetLifeInsurance1"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "63. Seguro de vida MET LIFE I"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "MetLifeInsurance2"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "64. Seguro de vida MET LIFE II"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "MortgageCredit"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "56. Crédito hipotecario"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "MortgageInsurance"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "55. Seguro hipotecario"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "NewEmployees"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Alta de empleados"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "PayrollRevision"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Revisión de nóminas"
			Select Case CInt(Request.Cookies("SIAP_SectionID"))
				Case 1
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
				Case 2
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
				Case 4
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
				Case 7
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
				Case Else
					aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			End Select
		Case "PayrollReviews"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Revisiones salariales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "PersonalLoan"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "60. Préstamo personal"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "ProfessionalRisk"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Carga de matriz de riesgos profesionales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "Third"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cargas de terceros institucionales"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Case "ThirdUploadMovements"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Archivos de terceros cargados"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "UpdateEmployeesData"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cambios a la información de los empleados"
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case "ProcessForSar"
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Ejercicio Bimestral del SAR."
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case Else
			Response.Redirect "Main.asp"
	End Select
	bWaitMessage = True
	%>
	<HTML>
		<HEAD>
			<!-- #include file="_JavaScript.asp" -->
		</HEAD>
		<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
			<%If (StrComp(CStr(oRequest("sAction").Item), "EmployeesMovements", vbBinaryCompare) <> 0) Or (StrComp(CStr(oRequest("sAction").Item), "EmployeesAssignNumber", vbBinaryCompare) <> 0) Or (StrComp(CStr(oRequest("sAction").Item), "Jobs", vbBinaryCompare) <> 0) Then
				Select Case sAction
					Case "Absences"
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Exportar a Excel las incidencias en proceso del empleado",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=Absences&Excel=1&Active=0&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True),_
							Array("Exportar a Excel las incidencias activas del empleado",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=Absences&Excel=1&Active=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True),_
							Array("Incidencias en proceso del empleado (corto)",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=Absences&ShortFormat=1&Excel=1&Active=0&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True),_
							Array("Incidencias activas del empleado (corto)",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=Absences&ShortFormat=1&Excel=1&Active=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
						)
					Case "AlimonyTypes"
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Exportar a Excel los tipos de pensiones registradas",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=AlimonyTypes&Excel=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
						)
					Case "ConceptsValues"
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Exportar a Excel los tabuladores en proceso",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=ConceptsValues&Excel=1&Active=0&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True And (CInt(Request.Cookies("SIAP_SubSectionID")) <> 32)),_
							Array("Exportar a Excel los tabuladores activos",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=ConceptsValues&Excel=1&Active=1&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
						)
					Case "MedicalAreas"
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Exportar a Excel los registros Unimed",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=MedicalAreas&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
						)
					Case "PayrollRevision"
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Exportar a Excel los registros de revisión del empleado",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Action=PayrollRevision&Excel=1" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
						)
					Case Else
						Select Case lReasonID
							Case EMPLOYEES_BANK_ACCOUNTS
								aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
									Array("Exportar a Excel los registros de cuentas bancarias de los empleados",_
										  "",_
										  "", "javascript: OpenNewWindow('Export.asp?Action=EmployeesMovements&Excel=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
								)
							Case EMPLOYEES_ADD_BENEFICIARIES
								aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
									Array("Exportar a Excel los registros de beneficiarios",_
										  "",_
										  "", "javascript: OpenNewWindow('Export.asp?Action=EmployeeBeneficiaries&Excel=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
								)
							Case Else
								Select Case sAction
									Case "ApplyAbsences"
										aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
											Array("Exportar a Excel las incidencias en proceso",_
												  "",_
												  "", "javascript: OpenNewWindow('Export.asp?Action=ApplyAbsences&Excel=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
										)
									Case Else
										aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
											Array("Exportar a Excel los movimientos en proceso",_
												  "",_
												  "", "javascript: OpenNewWindow('Export.asp?Action=EmployeesMovements&Excel=1&ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
										)
								End Select
						End Select
				End Select
				aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 703
				aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
				aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 290
			End If%>
			<!-- #include file="_Header.asp" -->
			<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
			If Not B_ISSSTE Then
				Response.Write "<A HREF=""HumanResources.asp"">Personal</A> > <A HREF=""Employees.asp"">Empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
			Else
				Select Case sAction
					Case "PayrollRevision"
						If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 25 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivos</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "AlimonyTypes"
						If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
							Select Case lReasonID
								Case CREDITORS_TYPES
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=27"">Acreedores de los empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Case Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Pensión alimenticia</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							End Select
						Else
							Select Case lReasonID
								Case CREDITORS_TYPES
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=27"">Acreedores de los empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Case Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Pensión alimenticia</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							End Select
						End If
					Case "Third"
						Dim sCreditTypeShortName
						Select Case sThirdConcept
							Case "ISSSTE"
								sCreditTypeShortName = "ISSSTE. Préstamos"
							Case "FOVISSSTE_62", "FOVISSSTE_86", "FOVISSSTE_56", "FOVISSSTE_NF"
								sCreditTypeShortName = "FOVISSSTE, Crédito hipotecario"
							Case Else
								Call GetNameFromTableByShortName(oADODBConnection, "CreditTypes", sThirdConcept, "", "", sCreditTypeShortName, "")
						End Select
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <A HREF=""Main_ISSSTE.asp?SectionID=211"">Carga de discos de terceros</A> > <B>" & sCreditTypeShortName & "</B><BR /><BR />"
					Case "ApplyAbsences"
						If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Aplicación de Incidencias</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Aplicación de Incidencias</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Employees.asp"">Movimientos de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "Absences"
						If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Incidencias</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Incidencias</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Employees.asp"">Movimientos de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "ChildrenSchoolarships", "EmployeesAntiquities", "EmployeesAnualAward", "EmployeesCarLoan", "EmployeesExtraHours", "EmployeesChildren", "EmployeesFamilyDeath", "EmployeesGlasses", "EmployeeMonthAward", "EmployeesProfessionalDegree", "EmployeeSports", "EmployeesSundays"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e Incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
					Case "CreditFOVISSSTE", "MetLifeInsurance1", "MetLifeInsurance2", "MortgageCredit", "MortgageInsurance", "PersonalLoan"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Captura de terceros</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
					Case "EmployeesAssignNumber"
						If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > "
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > "
						End If
						If iStep = 3 Then
							Response.Write "<A HREF=""UploadInfo.asp?Action=EmployeesAssignNumber&ReasonID=0"">" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</A> > <B>Resultado de la carga por archivo<BR /><BR />"
						Else
							Response.Write "<B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "EmployeesAssignTemporalNumber"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > " & aHeaderComponent(S_TITLE_NAME_HEADER) & "<BR /><BR />"
					Case "EmployeesConcepts"
						If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						Else
							If CInt(Request.Cookies("SIAP_SubSectionID")) = 25 Then
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Else
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							End If
						End If
					Case "EmployeesMovements", "EmployeesAssignJob", "EmployeesChanges", "EmployeesLicenses", "EmployeesNew", "EmployeesForRisk", "EmployeesDrop", "ResumptionOfWork"
						Select Case lReasonID
							Case -91
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>" & "Aplicación de registros cargados por cada archivo" & "</B><BR /><BR />"
							Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=712"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=20"">SI. Seguro de separación y AE. Seguro adicional de separación individualizado</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_FOR_RISK
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=712"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=18"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case -58
								If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									If CInt(Request.Cookies("SIAP_SubSectionID")) = 25 Then
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Registro de reclamos</B><BR /><BR />"
									Else
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>Registro de reclamos</B><BR /><BR />"
									End If
								End If
							Case -89, EMPLOYEES_NON_EXCENT
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_EXCENT
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=721"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI ', CANCEL_EMPLOYEES_C04
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									'If lReasonID = CANCEL_EMPLOYEES_C04 Then
									'	Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=18"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									'Else
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=721"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									'End If
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
									'If lReasonID = CANCEL_EMPLOYEES_C04 Then
									'	Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=18"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									'Else
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=17"">SI. Seguro de separación y AE. Seguro adicional de separación individualizado</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									'End If
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=20"">SI. Seguro de separación y AE. Seguro adicional de separación individualizado</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_ANTIQUITIES, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_CONCEPT_16, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_CONCEPT_7S, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									If CInt(Request.Cookies("SIAP_SubSectionID")) = 721 Then
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=721"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									Else
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									End If
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									If CInt(Request.Cookies("SIAP_SubSectionID")) = 211 Then
										'Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									Else
										Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=22"">Prestaciones e incidencias</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									End If
								End If
							Case EMPLOYEES_BENEFICIARIES, EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_BENEFICIARIES_DEBIT, ALIMONY_TYPES
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Pensión alimenticia</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case EMPLOYEES_CREDITORS, CREDITORS_TYPES
								If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=27"">Acreedores de los empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=27"">Acreedores de los empleados</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_LICENSES, -91
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case 54
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Jobs.asp"">Administración de Plazas</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Jobs.asp"">Administración de Plazas</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_THIRD_CONCEPT
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case EMPLOYEES_BANK_ACCOUNTS
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Registro de Cuentas Bancarias</B><BR /><BR />"
								ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Registro de Cuentas Bancarias</B><BR /><BR />"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								End If
							Case EMPLOYEES_GRADE
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								'If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
								'	Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Calificación de empleados</B><BR /><BR />"
								'ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
								'	Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Calificación de empleados</B><BR /><BR />"
								'Else
								'	Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								'End If
							Case 58
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=18"">Administración de personal</A> > <A HREF=""UploadInfo.asp?Action=EmployeesAssignNumber&ReasonID=0"">Asignación de número de empleado</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case "ProcessForSar"
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Reportes</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
							Case Else
								If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=712"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
								Else
									Select Case CInt(Request.Cookies("SIAP_SubSectionID"))
										Case 2
											Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=16"">Reclamo de pago por ajustes y deducciones</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
										Case Else
											Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=18"">Administración de personal</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
									End Select
								End If
						End Select
					Case "FONAC"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Entrada del archivo de FOVISSSTE</B><BR /><BR />"
					Case "Jobs"
						If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > "
                        ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
                            Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > "
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > "
						End If
						If iStep = 3 Then
							Response.Write "<A HREF=""Jobs.asp"">Administración de Plazas</A> > <A HREF=""UploadInfo.asp?Action=Jobs&ReasonID=" & lReasonID & """>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</A> > <B>Resultado de la carga por archivo</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Jobs.asp"">Administración de Plazas</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "JobServices"
						If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Jobs.asp"">Administración de Plazas</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Jobs.asp"">Administración de Plazas</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "MedicalAreas"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
					Case "ConceptsValues"
						If CInt(Request.Cookies("SIAP_SubSectionID")) <> 32 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <A HREF=""Main_ISSSTE.asp?SectionID=33"">Registro de tabuladores</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <A HREF=""Main_ISSSTE.asp?SectionID=32"">Consulta de tabuladores</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "ThirdUploadMovements"
						'Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Carga de conceptos de terceros</B><BR /><BR />"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Aplicación de registros cargados por cada archivo</B><BR /><BR />"
					Case "ProcessForSar"
						If StrComp(oRequest("Load").Item,"PayrollSummary",vbBinaryCompare) = 0 Then
							aHeaderComponent(S_TITLE_NAME_HEADER) = "Cargar resumen de nóminas"
						ElseIf StrComp(oRequest("Load").Item,"BanamexCensus",vbBinaryCompare) = 0 Then
							aHeaderComponent(S_TITLE_NAME_HEADER) = "Cargar Padrón SAR"
						ElseIf StrComp(oRequest("Load").Item,"ConsarFile",vbBinaryCompare) = 0 Then
							aHeaderComponent(S_TITLE_NAME_HEADER) = "Cargar archivo de línea de captura"
						ElseIf StrComp(oRequest("Load").Item,"SarCensus",vbBinaryCompare) = 0 Then
							aHeaderComponent(S_TITLE_NAME_HEADER) = "Actualización masiva de padrón SAR"
						End If
						If iStep = 3 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio Bimestral del SAR</A> > <A HREF=""UploadInfo.asp?Action=ProcessForSar&Load=PayrollSummary"">" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</A> <B>Resultado de la carga por archivo</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio Bimestral del SAR</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
						End If
					Case "ProfessionalRisk"
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=29"">Matriz de riesgos profesionales</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
					Case Else
						Response.Write "<BR /><BR />"
				End Select
			End If

			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Error al registrar la información", sErrorDescription)
			End If
			If iStep <= 1 Then
				Dim sAltDescription
				Dim sDescription
				Select Case sAction
					Case "EmployeesMovements"
						Select Case lReasonID
							Case EMPLOYEES_BANK_ACCOUNTS
								sAltDescription = "Cuenta bancaria"
								sDescription = "Registre una cuenta bancaria a un empleado diferente."
							Case EMPLOYEES_ADD_BENEFICIARIES
								sAltDescription = "Beneficiario de pensión alimenticia"
								sDescription = "Registre un(a) beneficiario(a) de pensión alimenticia a un empleado diferente."
							Case EMPLOYEES_CREDITORS
								sAltDescription = "Acreedores"
								sDescription = "Registre un(a) acreedor(a) de adeudos a un empleado diferente."
							Case -96,-75,-64,1,2,3,4,5,6,7,8,10,12,13,14,17,18,21,26,28,29,30,31,32,33,34,37,38,39,40,41,43,44,45,46,47,48,50,51,53,57,58,62,63,66,68,78,79,80,81,101,102,103,104,105,106
								sAltDescription = "Movimientos de personal"
								sDescription = "Registre el movimiento a un empleado diferente."
							Case Else
								sAltDescription = "Prestación"
								sDescription = "Registre la prestación a un empleado diferente."
						End Select
					Case "Absences"
						sAltDescription = "Incidencias"
						sDescription = "Registre incidencias a un empleado diferente."
				End Select
				If Len(oRequest("EmployeeNumber").Item) > 0 Then
					aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(aEmployeeComponent(S_NUMBER_EMPLOYEE))
					lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						If VerifyRequerimentsForEmployeesConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
							lErrorNumber = DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
						Else
							lErrorNumber = -1
							Select Case sAction
								Case "EmployeesMovements"
									Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
								Case "Absences"
									Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
							End Select
						End If
					Else
						lErrorNumber = -1
						Select Case sAction
							Case "EmployeesMovements"
								Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
							Case "Absences"
								Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
						End Select
					End If
				ElseIf Len(oRequest("EmployeeID").Item) > 0 Then
					Select Case sAction
						Case "PayrollRevision"
							lErrorNumber = DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
						Case "Absences"
							If Len(oRequest("AbsenceChange").Item) > 0 Then
								lErrorNumber = DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
							Else
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									If VerifyRequerimentsForEmployeesAbsences(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
										lErrorNumber = DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
									Else
										lErrorNumber = -1
										Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
									End If
								Else
									Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
								End If
							End If
						Case "EmployeesMovements"
							If VerifyRequerimentsForEmployeesConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
								lErrorNumber = DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
							Else
								Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "UploadInfo.asp", sAction, 10, lReasonID, sAltDescription, sDescription, sErrorDescription)
							End If
					End Select
				Else
					If InStr(1, sAction, "AlimonyTypes") > 0 Then
						Call DisplayAlimonyTypesForm(oRequest, oADODBConnection, GetASPFileName(""), lReasonID, aEmployeeComponent, sErrorDescription)	
					Else
						Call DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
					End If
				End If
				If Len(oRequest("Success").Item) > 0 Then
					If CInt(oRequest("Success").Item) = 1 Then
						Select Case sAction
							Case "Absences"
								'Call DisplayErrorMessage("Confirmación", "La operación con la incidencia " & sAbsenceShortName & " fué ejecutada exitosamente.")
							Case Else
								Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente." & CStr(oRequest("ErrorDescription").Item))
						End Select
					Else
						Select Case sAction
							Case "Absences"				
								'Call DisplayErrorMessage("Error al realizar la operación con la incidencia " & sAbsenceShortName, CStr(oRequest("ErrorDescription").Item))
							Case Else
								Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("ErrorDescription").Item))
						End Select
					End If
				End If
			Else
				Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 1. </B>Introduzca el archivo a utilizar.<BR /><BR />"
				Select Case sAction
					Case "Third"
						Select Case iStep
							Case 2
								If bUploadFile Then
									Call UploadThirdFile(sThirdConcept, sAction, oADODBConnection, sUploadFile, sOriginalFileName, sErrorDescription)
									Call  DeleteFile(sUploadFile, sUploadFileError)
								Else
									Call UploadThirdFile(sThirdConcept, sAction, oADODBConnection, sFileName, sOriginalFileName, sErrorDescription)
								End If
						End Select
					Case "ChildrenSchoolarships"
						Select Case iStep
							Case 2
								Call DisplayChildrenSchoolarshipsColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadChildrenSchoolarshipsFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Las becas de los hijos de los empleados fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar las becas de los hijos de los empleados.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "ConceptsValues"
						Select Case iStep
							Case 2
								Call DisplayConceptsValuesColumns(sFileName, lEmployeeTypeID, False, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadConceptsValuesFile(oADODBConnection, sFileName, False, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar los tabuladores.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "CreditFOVISSSTE"
					Case "DocumentsForLicenses"
						Select Case iStep
							Case 2
								Call DisplayDocumentsForLicensesColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Identifique las columnas del archivo.<BR /><BR />"
								lErrorNumber = UploadDocumentsForLicensesFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar a los empleados", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "Absences", "EmployeesAbsences"
						Select Case iStep
							Case 2
								Call DisplayEmployeesAbsencesColumns(lReasonID, sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Identifique las columnas del archivo.<BR /><BR />"
								lErrorNumber = UploadEmployeesAbsencesFile(oADODBConnection, lReasonID, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Las incidencias fueron registradas con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar las incidencias", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "EmployeesAccounts"
					Case "EmployeesMovements"
						Select Case iStep
							Case 2
								Select Case lReasonID
									Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
										Call DisplayDocumentsForLicensesColumns(sFileName, sErrorDescription)
									Case -58
										Call DisplayEmployeesAdjustmentsColumns(sFileName, sErrorDescription)
									'Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
									'	Call DisplayEmployeesSafeSeparationColumns(lReasonID, sAction, sFileName, sErrorDescription)
									Case -89, EMPLOYEES_FOR_RISK, EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, EMPLOYEES_ANTIQUITIES, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CONCEPT_08, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_EFFICIENCY_AWARD, EMPLOYEES_GRADE, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
										Call DisplayEmployeesFeaturesColumns(lReasonID, sAction, sFileName, sErrorDescription)
									Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
										If bUploadFile Then
											Call DisplayEmployeesAbsencesColumns(lReasonID, sUploadFile, sErrorDescription)
										Else
											Call DisplayEmployeesAbsencesColumns(lReasonID, sFileName, sErrorDescription)
										End If
									Case EMPLOYEES_BENEFICIARIES_DEBIT
										Call DisplayEmployeesBeneficiariesDebitColumns(lReasonID, sFileName, sErrorDescription)
									Case EMPLOYEES_BANK_ACCOUNTS
										Call DisplayEmployeesBankAccountColumns(lReasonID, sFileName, sErrorDescription)
									Case Else
										Call DisplayRegisterEmployeesColumns(sFileName, sAction, lReasonID, sErrorDescription)
								End Select
							Case 3
								Select Case lReasonID
									Case -58
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										lErrorNumber = UploadEmployeesAdjustmentsFile(oADODBConnection, sFileName, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
										Else
											Call DisplayErrorMessage("Error al registrar los reclamos de pago por ajustes y deducciones.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Identifique las columnas del archivo.<BR /><BR />"
										lErrorNumber = UploadDocumentsForLicensesFile(oADODBConnection, sFileName, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
										Else
											Call DisplayErrorMessage("Error al registrar a los empleados", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									'Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
									'	Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
									'	lErrorNumber = UploadEmployeesSafeSeparationFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
									'	Response.Write "<BR />"
									'	If lErrorNumber = 0 Then
									'		Call DisplayErrorMessage("Confirmación", "Los registros fueron registrados con éxito.")
									'	Else
									'		Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
									'		lErrorNumber = 0
									'		sErrorDescription = ""
									'	End If
									Case -89, EMPLOYEES_FOR_RISK, EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, EMPLOYEES_ANTIQUITIES, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CONCEPT_08, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_EFFICIENCY_AWARD, EMPLOYEES_GRADE, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										lErrorNumber = UploadEmployeesFeaturesFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
										Else
											Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										If bUploadFile Then
											lErrorNumber = UploadEmployeesAbsencesFile(oADODBConnection, lReasonID, sUploadFile, sErrorDescription)
											Call  DeleteFile(sUploadFile, sUploadFileError)
										Else
											lErrorNumber = UploadEmployeesAbsencesFile(oADODBConnection, lReasonID, sFileName, sErrorDescription)
										End If
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
										Else
											Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Case EMPLOYEES_BENEFICIARIES_DEBIT
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										lErrorNumber = UploadEmployeesBeneficiariesDebitFile(lReasonID, sAction, oADODBConnection, sFileName, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "La información fue resgistrada exitosamente.")
										Else
											Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Case EMPLOYEES_BANK_ACCOUNTS
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										lErrorNumber = UploadEmployeesBankAccountFile(oADODBConnection, sFileName, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "Los registros fueron realizados con éxito.")
										Else
											Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Case Else
										sMessage = "Alta de empleados"
										Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
										lErrorNumber = UploadRegisterEmployeesFile(oADODBConnection, sFileName, "EmployeesMovements", lReasonID, sErrorDescription)
										Response.Write "<BR />"
										If lErrorNumber = 0 Then
											Call DisplayErrorMessage("Confirmación", "La información de los empleados fue registrada correctamente en el sistema.")
										Else
											Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
								End Select
						End Select
					Case "EmployeesAnualAward"
					Case "EmployeesAssignNumber"
						Select Case iStep
							Case 2
								Call DisplayEmployeesAssignNumberColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadEmployeesAssignNumberFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Se han generado los número de empleado con éxito.")
								Else
									Call DisplayErrorMessage("Error al generar los números de empleados.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "EmployeesChildren"
						Select Case iStep
							Case 2
								Call DisplayEmployeesChildrenColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadEmployeesChildrenFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los hijos de los empleados fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar a los hijos de los empleados.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "EmployeesCarLoan"
					Case "EmployeesExtraHours"
						Select Case iStep
							Case 2
								Call DisplayEmployeesExtraHoursColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadEmployeesExtraHoursFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los montos a pagar a los empleados por concepto 09 fueron registradas con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar el concepto 09.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "EmployeesNew"
						Select Case iStep
							Case 2
								Call DisplayRegisterEmployeesColumns(sFileName, "EmployeesNew2", 12, sErrorDescription)
							Case 3
								sMessage = "Alta de empleados"
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadRegisterEmployeesFile(oADODBConnection, sFileName, "EmployeesNew2", lReasonID, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar la información de los empleados.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "EmployeesAssignJob", "EmployeesDrop"
						Select Case iStep
							Case 2
								Call DisplayRegisterEmployeesColumns(sFileName, sAction, lReasonID, sErrorDescription)
							Case 3
								Select Case sAction
									Case "EmployeesAssignJob" sMessage = "Cambio de plaza"
									Case "EmployeesAssignJob" sMessage = "Baja de empleados"
								End Select
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadRegisterEmployeesFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar la información de los empleados.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
									'Call DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, "0", lReasonID, aEmployeeComponent, sErrorDescription)
								End If
						End Select
					Case "EmployeesFONAC"
					Case "EmployeesInactive"
					Case "EmployeesLicenses"
					Case "EmployeesNightShifts"
					Case "EmployeesResumptions"
					Case "EmployeesSAR"
					Case "FONAC"
						Select Case iStep
							Case 2
								lErrorNumber = UploadFONACFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If Len(sErrorDescription) = 0 Then
									Call DisplayErrorMessage("Confirmación", "El archivo de FOVISSSTE fue registrado con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar el archivo de FOVISSSTE", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "Jobs"
						Select Case iStep
							Case 2
								Call DisplayJobsColumns(sFileName, sAction, lReasonID, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadJobsFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar la información de las plazas.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "JobServices"
						Select Case iStep
							Case 2
								Call DisplayRegisterEmployeesColumns(sFileName, sAction, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Identifique las columnas del archivo.<BR /><BR />"
								lErrorNumber = UploadRegisterEmployeesFile(oADODBConnection, sFileName, sAction, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los cambios de servico a la plaza fueron registradas con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar los cambios de servicios a las plazas", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "MedicalAreas"
						Select Case iStep
							Case 2
								Call DisplayMedicalAreasColumns(sFileName, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadMedicalAreasFile(oADODBConnection, sFileName, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "El archivo con la información UNIMED fue registrado con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar el archivo de la información UNIMED.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "ProcessForSar"
						Select Case iStep
							Case 2
								Call DiplaySarProcessColumns(sFileName, sAction, lReasonID, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								If StrComp(oRequest("Load").Item, "PayrollSummary", vbBinaryCompare) = 0 Then
									lErrorNumber = UploadHistNomsarFile(oADODBConnection, sFileName, sErrorDescription)
								ElseIf (StrComp(oRequest("Load").Item, "BanamexCensus", vbBinaryCompare) = 0) Or _
										(StrComp(oRequest("Load").Item, "SarCensus", vbBinaryCompare) = 0) Then
									lErrorNumber = UploadBanamexCensus(oADODBConnection, sFileName, sErrorDescription)
									If StrComp(oRequest("Load").Item,"SarCensus",VbBinaryCompare) = 0 Then
										If lErrorNumber = 0 Then
											lErrorNumber = CompareSarCensus(oRequest, oADODBConnection, sErrorDescription)
										End If
									End If
								ElseIf (StrComp(oRequest("Load").Item, "ConsarFile", vbBinaryCompare) = 0) Then
									lErrorNumber = UploadConsarFile(oADODBConnection, sFileName, sErrorDescription)
								End If
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "La información fue registrada correctamente.")
								Else
									Call DisplayErrorMessage("Error al registrar la información.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "ProfessionalRisk"
						Select Case iStep
							Case 2
								Call DisplayProfessionalRiskColumns(sFileName, sAction, lReasonID, sErrorDescription)
							Case 3
								Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
								lErrorNumber = UploadProfessionalRiskFile(oADODBConnection, sFileName, sAction, lReasonID, sErrorDescription)
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Call DisplayErrorMessage("Confirmación", "Los registros fueron registrados con éxito.")
								Else
									Call DisplayErrorMessage("Error al registrar la información de las plazas.", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
						End Select
					Case "MetLifeInsurance1"
					Case "MetLifeInsurance2"
					Case "MortgageCredit"
					Case "MortgageInsurance"
					Case "NewEmployees"
					Case "PayrollReviews"
					Case "PersonalLoan"
				End Select
			End If
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Error al realizar la operación", sErrorDescription)
				Response.Write "<BR />"
				lErrorNumber = 0
				sErrorDescription = ""
			End If%>
			<!-- #include file="_Footer.asp" -->
		</BODY>
	</HTML>
<%End If%>