<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/PayrollComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<%
Dim sAction
Dim bShowForm
Dim bShowTable
Dim lPayrollID
Dim iPayrollStatus
Dim sNames
Dim iSelectedTab
Dim lRecordID
Dim lRecordID2
Dim sAuthorize
Dim lSuccess
Dim lPositionID
Dim iStep
Dim sFileName
Dim lEmployeeTypeID
Dim sRequiredFields

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PAYROLL_PERMISSIONS) = N_PAYROLL_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_PAYROLL_PERMISSIONS
	End If
End If

sAction = oRequest("Action").Item
lRecordID = oRequest("RecordID").Item
lRecordID2 = oRequest("RecordID2").Item
sAuthorize = oRequest("Authorize").Item
bShowTable = False

If Len(oRequest("EmployeeTypeID").Item)>0 Then
	lEmployeeTypeID = CLng(oRequest("EmployeeTypeID").Item)
Else
	lEmployeeTypeID = -1
End If

iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)
sFileName = Server.MapPath(UPLOADED_PHYSICAL_PATH & sAction & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt")
If Len(oRequest("RawData").Item) > 0 Then
	lErrorNumber = SaveTextToFile(sFileName, oRequest("RawData").Item, sErrorDescription)
	If lErrorNumber = 0 Then
		Select Case sAction
			Case "ConceptsValues"
				Response.Redirect "Payroll.asp?Action=" & sAction & "&EmployeeTypeID=" & lEmployeeTypeID & "&Step=" & iStep
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

If (InStr(1, sAction , "Concepts", vbBinaryCompare) <> 1) Then aConceptComponent(N_START_DATE_CONCEPT) = ""
iSelectedTab = -1
If Len(oRequest("Tab").Item) > 0 Then
	iSelectedTab = CInt(oRequest("Tab").Item)
ElseIf Len(oRequest("EmployeeTypeID").Item) > 0 Then
	iSelectedTab = CInt(oRequest("EmployeeTypeID").Item)
End If
If Len(oRequest("Success").Item) > 0 Then lSuccess = CInt(oRequest("Success").Item)
If Len(oRequest("PositionID").Item) > 0 Then
	lPositionID = CLng(oRequest("PositionID").Item)
Else
	lPositionID = -1
End If

Call InitializeConceptComponent(oRequest, aConceptComponent)
Call InitializePayrollComponent(oRequest, aPayrollComponent)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
Select Case sAction
	Case "ConceptsFile"
		lErrorNumber = GetLastPayrollStatus(oADODBConnection, lPayrollID, iPayrollStatus, sErrorDescription)
		lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, -1, lPayrollID, sErrorDescription)
		bShowTable = True
		If B_ISSSTE Then
			If iGlobalSectionID = 4 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Tabuladores"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
		End If
	Case "Concepts"
		If Len(oRequest("Add").Item) > 0 Then
			lErrorNumber = AddConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=Concepts" & "&New=1&Success=1"
			Else
				Response.Redirect "Payroll.asp?Action=Concepts" & "&New=1&Success=0&sErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			lErrorNumber = ModifyConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=1"
			Else
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=0&sErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Delete").Item) > 0 Then
			lErrorNumber = RemoveConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=1"
			Else
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=0&sErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Apply").Item) > 0 Then
			lErrorNumber = SetActiveForConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=1"
			Else
				Response.Redirect "Payroll.asp?Action=Concepts" & "&Success=0&sErrorDescription=" & sErrorDescription
			End If
			'Redim aConceptComponent(N_CONCEPT_COMPONENT_SIZE)
			'aConceptComponent(N_ID_CONCEPT) = -1
			'bShowForm = False
		End If
		bShowTable = True
		If B_ISSSTE Then
			If iGlobalSectionID = 4 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Tabuladores"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
		End If
	Case "ConceptsValues"
		Dim oRecordset
		Dim iErrorCount
		Dim iSuccessCount
		iErrorCount = 0
		iSuccessCount = 0
		If Len(oRequest("Add").Item) > 0 Then
			lErrorNumber = AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab & "&Success=1"
			Else
				Response.Redirect "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab & "&Success=0&sErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			If Len(oRequest("bFull").Item) > 0 Then 
				aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
			End If
			lErrorNumber = ModifyConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
			aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
			aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
			aConceptComponent(N_WORKING_HOURS_CONCEPT) = -1
			aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
			aConceptComponent(N_HAS_CHILDREN_CONCEPT) = -1
			aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0
			aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1
			aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
			aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
			aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = 0
			aConceptComponent(N_CURRENCY_ID_CONCEPT) = 0
			aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = 1
			aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = 3
			aConceptComponent(N_APPLIES_ID_CONCEPT) = -1
			aConceptComponent(N_START_USER_ID_CONCEPT) = aLoginComponent(N_USER_ID_LOGIN)
			aConceptComponent(N_END_USER_ID_CONCEPT) = -1
			'If Len(oRequest("Perceptions").Item) > 0 Then
				If lErrorNumber = 0 Then
					Response.Redirect "Payroll.asp?Action=ConceptsValues&Success=1&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)
				Else
					Response.Redirect "Payroll.asp?Action=ConceptsValues&Success=0&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)
				End If
			'End If
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			lErrorNumber = RemoveConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			aConceptComponent(N_RECORD_ID_CONCEPT) = -1
			Response.Redirect "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab
		ElseIf Len(oRequest("Authorize").Item) > 0 Then
			lErrorNumber = AuthorizeConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			aConceptComponent(N_RECORD_ID_CONCEPT) = -1
			Response.Redirect "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab 
		ElseIf Len(oRequest("Apply").Item) Then
			aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
			lErrorNumber = SetActiveForConceptsValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Payroll.asp?Action=ConceptsValues&Success=1&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab
			Else
				Response.Redirect "Payroll.asp?Action=ConceptsValues&Success=0&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab
			End If
		ElseIf (Len(oRequest("AuthorizationFile").Item) > 0) Then
			lErrorNumber = AddConceptsValuesFileSP(oRequest, oADODBConnection, oRequest("sQuery").Item, aConceptComponent, sErrorDescription)
			sError = sErrorDescription
			If lErrorNumber = 0 Then
				sError = sError & "El puesto para guardias y suplencias se registró exitosamente<BR />"
			Else
				sError = sError & "Error al registrar el puesto para guardias y suplencias<BR />"
			End If
		End If
		bShowForm = True
		bShowTable = True
		If B_ISSSTE Then
			If iGlobalSectionID = 4 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Tabuladores"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Conceptos de pago"
		End If
	Case "EmployeeTypes"
		bShowTable = True
		If B_ISSSTE Then
			If iGlobalSectionID = 4 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Tipos de tabuladores"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Tipos de tabuladores"
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Tipos de tabuladores"
		End If
	Case Else
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Nómina"
		lErrorNumber = GetLastPayrollStatus(oADODBConnection, lPayrollID, iPayrollStatus, sErrorDescription)
		If Len(oRequest("Add").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Crear una nueva nómina"
			lErrorNumber = AddPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
			'If lErrorNumber = 0 Then
			'	bShowTable = True
			'	lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 1, aPayrollComponent, sErrorDescription)
			'End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Modificar nómina"
			lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
			bShowTable = True
		ElseIf Len(oRequest("CalculatePayroll").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Prenómina"
			If Len(Application.Contents("SIAP_CalculatePayroll")) = 0 Then
				Application.Contents("SIAP_CalculatePayroll") = aLoginComponent(N_USER_ID_LOGIN) & LIST_SEPARATOR & aPayrollComponent(N_ID_PAYROLL) & LIST_SEPARATOR & GetSerialNumberForDate("")
				lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 2, aPayrollComponent, sErrorDescription)
			End If
			If lErrorNumber =  0 Then
				lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				If lErrorNumber =  0 Then
					aPayrollComponent(N_CLOSED_PAYROLL) = 2
					lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				End If
			End If
		ElseIf Len(oRequest("DoClose").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cerrar una nómina"
			If StrComp(oRequest("DoMessages").Item, "1", vbBinaryCompare) = 0 Then
				If Len(Application.Contents("SIAP_CalculatePayroll")) = 0 Then
					Application.Contents("SIAP_CalculatePayroll") = aLoginComponent(N_USER_ID_LOGIN) & LIST_SEPARATOR & aPayrollComponent(N_ID_PAYROLL) & LIST_SEPARATOR & GetSerialNumberForDate("")
					lErrorNumber = InsertEmployeesChangesLKP(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				End If
			ElseIf StrComp(oRequest("DoMessages").Item, "2", vbBinaryCompare) = 0 Then
				Call GetPayroll(oRequest, oADODBConnection, aPayrollComponent, "")
				aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))

				lErrorNumber = InsertPaymentMessages(oADODBConnection, aPayrollComponent, sErrorDescription)
			Else
				lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 3, aPayrollComponent, sErrorDescription)
				If lErrorNumber =  0 Then
					lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
					If lErrorNumber =  0 Then
						aPayrollComponent(N_CLOSED_PAYROLL) = 1
						lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
					End If
				End If
			End If
			bShowTable = True
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Borrar una nómina"
			lErrorNumber = RemovePayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
			If lErrorNumber = 0 Then aPayrollComponent(N_ID_PAYROLL) = -1
			bShowTable = True
		ElseIf (Len(sAction) > 0) Then
			'bShowTable = (StrComp(sAction, "RetroactivePayroll", vbBinaryCompare) <> 0)
			bShowTable = (InStr(1, ",ModifyPayroll,", sAction, vbBinaryCompare) = 0)
		End If
End Select

bWaitMessage = True
Response.Cookies("SoS_SectionID") = 194
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If (iGlobalSectionID = 3) Or (iGlobalSectionID = 4) Then
			Select Case sAction
				Case "Concepts", "ConceptsFile", "ConceptsValues"
					If iGlobalSectionID = 4 Then
						aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Agregar un nuevo concepto",_
								  "",_
								  "", "Payroll.asp?Action=Concepts&New=1", (StrComp(sAction, "ConceptsValues", vbBinaryCompare) <> 0)),_
							Array("Agregar un nuevo monto",_
								  "",_
								  "", "Payroll.asp?Action=ConceptsValues&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptsValues", vbBinaryCompare) = 0)),_
							Array("Exportar a Excel",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptsValues", vbBinaryCompare) = 0)),_
							Array("Exportar histórico a Excel",_
								  "",_
								  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&HistoryList=1&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptsValues", vbBinaryCompare) = 0)),_
							Array("Generar archivos de cálculo",_
								  "",_
								  "", "Payroll.asp?Action=ConceptsFile", False)_
						)
					End If
				'Case "ConceptsValues"
				'	aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				'		Array("Agregar un nuevo tabulador",_
				'			  "",_
				'			  "", "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) <> -1)),_
				'		Array("Exportar a Excel",_
				'			  "",_
				'			  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) <> -1)),_
				'		Array("Generar archivos de cálculo",_
				'			  "",_
				'			  "", "Payroll.asp?Action=ConceptsFile", True)_
				'	)
			End Select
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		Else
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo tabulador",_
					  "",_
					  "", "Payroll.asp?Action=ConceptsValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) <> -1)),_
				Array("Generación de tabuladores para el ISSSTE",_
					  "",_
					  "", "Payroll.asp?Action=Concepts&UploadConcepts=1", True And False),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) <> -1))_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 713
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 280
		End If%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		Select Case sAction
			Case "AddPayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Crear una nueva nómina</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Nómina nueva</B>"
				End If
			Case "AddSpecialPayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Nóminas especiales</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Nómina nueva</B>"
				End If
			Case "ClosePayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Cerrar nómina</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Cerrar nómina</B>"
				End If
			Case "Concepts", "ConceptsFile"
				If B_ISSSTE Then
					If iGlobalSectionID = 4 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Conceptos de pago</B>"
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <B>Tabuladores</B>"
					End If
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Conceptos de pago</B>"
				End If
			Case "ConceptsValues"
				If Len(oRequest("Perceptions").Item) > 0 Then
					If iSelectedTab >= 0 Then
						Call GetNameFromTable(oADODBConnection, "EmployeeTypes", iSelectedTab, "", "", sNames, "")
					Else
						Call GetNameFromTable(oADODBConnection, "FullConcepts", oRequest("ConceptID").Item, "", "", sNames, "")
					End If
				Else
					Call GetNameFromTable(oADODBConnection, "FullConcepts", oRequest("ConceptID").Item, "", "", sNames, "")
				End If
				If B_ISSSTE Then
					If iGlobalSectionID = 4 Then
						Call GetNameFromTable(oADODBConnection, "FullConcepts", oRequest("ConceptID").Item, "", "", sNames, "")
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Payroll.asp?Action=Concepts"">Conceptos de pago</A> > <B>" & sNames & "</B>"
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <A HREF=""Payroll.asp?Action=EmployeeTypes"">Tabuladores</A> > <B>" & sNames & "</B>"
					End If
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <A HREF=""Payroll.asp?Action=Concepts"">Conceptos de pago</A> > <B>" & sNames & "</B>"
				End If
			Case "DeletePayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Borrar nueva</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Borrar nómina</B>"
				End If
			Case "EmployeeTypes"
				If B_ISSSTE Then
					If iGlobalSectionID = 4 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Conceptos de pago</B>"
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > <B>Tabuladores</B>"
					End If
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Conceptos de pago</B>"
				End If
			Case "ModifyPayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Prenómina</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Generar prenómina</B>"
				End If
			Case "RetroactivePayroll"
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Pagos retroactivos</B>"
				Else
					Response.Write "<A HREF=""Payroll.asp"">Nómina</A> > <B>Pagos retroactivos</B>"
				End If
			Case "UpdatePayroll"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Modificar nómina</B>"
			Case Else
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Nómina</B>"
				Else
					Response.Write "<B>Nómina</B>"
				End If
		End Select
		Response.Write "<BR /><BR />"

		If iStep <= 1 Then
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				Response.Write "<BR />"
			End If
			If StrComp(sAction, "ConceptsValues", vbBinaryCompare) = 0 Then
				If ((iGlobalSectionID <> 3) And (Len(oRequest("Perceptions").Item) > 0) Or (iGlobalSectionID = 4)) Then
					lErrorNumber = DisplayConceptValuesTabs(oRequest, aConceptComponent(N_ID_CONCEPT), iSelectedTab, sErrorDescription)
				End If
			End If
			If bShowTable Then
				Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
					Select Case sAction
						Case "Concepts", "ConceptsFile"
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['ConceptInfoFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros activos para el tipo de tabulador</FONT></A><BR /><BR />"
								Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
									Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Payroll.asp"" METHOD=""POST"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & oRequest("ConceptID").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & oRequest("StartDate").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><FONT FACE=""Arial"" SIZE=""2""> Mostrar los conceptos con clave: </FONT>"
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptShortName"" ID=""ConceptShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aConceptComponent(S_SHORT_NAME_CONCEPT)) & """ CLASS=""TextFields"" />"
										Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" VALUE=""Consultar registros"" CLASS=""Buttons""><BR />"
										'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""800"" HEIGHT=""1"" /><BR />"
									Response.Write "</FORM>"
									If Len(aConceptComponent(S_SHORT_NAME_CONCEPT)) > 0 Then aConceptComponent(S_QUERY_CONDITION_CONCEPT) = aConceptComponent(S_QUERY_CONDITION_CONCEPT) & " And (Concepts.ConceptShortName like '%" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "%')"
									Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
										Response.Write "<TR><TD>" & vbNewLine
											aConceptComponent(N_STATUS_ID_CONCEPT) = 1
											If iGlobalSectionID <> 4 Then aConceptComponent(S_QUERY_CONDITION_CONCEPT) = " And (Concepts.ConceptID In (1,3,14,38,39,49,89))"
											lErrorNumber = DisplayConceptsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (iGlobalSectionID = 4), aConceptComponent, sErrorDescription)
											If lErrorNumber <> 0 Then
												Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
												lErrorNumber = 0
												sErrorDescription = ""
											End If
										Response.Write "</TD></TR>" & vbNewLine
									Response.Write "</TABLE>" & vbNewLine
								Response.Write "</DIV>"
								Response.Write "<BR />"
								Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros que estan en proceso para el tipo de tabulador</FONT></A><BR /><BR />"
								Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
									Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
										Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
											Response.Write "<TR>"
												If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<TR><TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/></TD></TR>"
											Response.Write "</TR>"
										Response.Write "</TABLE>"
										Response.Write "<BR />"
										aConceptComponent(N_STATUS_ID_CONCEPT) = 0
										aConceptComponent(S_QUERY_CONDITION_CONCEPT) = ""
										lErrorNumber = DisplayConceptsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (iGlobalSectionID = 4), aConceptComponent, sErrorDescription)
										If lErrorNumber <> 0 Then
											Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Response.Write "</FORM>"
								Response.Write "</DIV>"
							Response.Write "</TD>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						Case "ConceptsValues"
							If iGlobalSectionID = 4 Then
								Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
									Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['ConceptInfoFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros activos para el tipo de tabulador</FONT></A><BR /><BR />"
									Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
										Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Payroll.asp"" METHOD=""POST"">"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & oRequest("ConceptID").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & oRequest("StartDate").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><FONT FACE=""Arial"" SIZE=""2""> Mostrar los tabuladores para el puesto: </FONT>"
											Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
											Response.Write "<TR>"
											If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1 Then
												Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
													Response.Write "<OPTION VALUE="""">Todos</OPTION>"
													Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Cia:' As Temp, CompanyID", "(PositionID>-1)", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
												Response.Write "</SELECT></TD>"
											ElseIf aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 1 Then
												Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
													Response.Write "<OPTION VALUE="""">Todos</OPTION>"
													Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions, GroupGradeLevels", "PositionID", "PositionShortName, PositionName, 'GGN:' As Temp1, GroupGradeLevelShortName, 'Clas:' As Temp2, ClassificationID, 'Int:' As Temp3, IntegrationID, 'Cia:' As Temp, CompanyID", "(Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
												Response.Write "</SELECT></TD>"
											Else
												Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
													Response.Write "<OPTION VALUE="""">Todos</OPTION>"
													Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Nivel:' As Temp, LevelID, 'Cia:' As Temp, CompanyID", "(PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT) & "," & aConceptComponent(N_LEVEL_ID_CONCEPT), "", sErrorDescription)
												Response.Write "</SELECT></TD>"
											End If
										Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" VALUE=""Consultar registros"" CLASS=""Buttons""></TD>"
									Response.Write "</FORM>"
									If aConceptComponent(N_POSITION_ID_CONCEPT) <> -1 Then aConceptComponent(S_QUERY_CONDITION_CONCEPT) = aConceptComponent(S_QUERY_CONDITION_CONCEPT) & " And (ConceptsValues.PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")"
									Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
										Response.Write "<TR><TD>" & vbNewLine
											aConceptComponent(N_STATUS_ID_CONCEPT) = 1
											lErrorNumber = DisplayConceptValuesTable(oRequest, oADODBConnection, iSelectedTab, False, sErrorDescription)
											If lErrorNumber <> 0 Then
												Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
												lErrorNumber = 0
												sErrorDescription = ""
											End If
										Response.Write "</TD></TR>" & vbNewLine
									Response.Write "</TABLE>" & vbNewLine
									Response.Write "</DIV>"
									Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
									Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
									Response.Write "<FORM NAME=""UploadInfoFrm"" ID=""UploadInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
										If InStr(1, sAction, "ConceptsValues") Then
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & lEmployeeTypeID & """ />"
										End If
										Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																			"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																			"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																			"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																			"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
																		"</OL></BLOCKQUOTE>")
										Response.Write "<BR />"
										sRequiredFields = "Tipo de tabulador*, Compañía*, Clave concepto de pago, Tipo de puesto*, Clave del puesto, Nivel*, Grupo grado nivel*, Estatus del empleado*, Estatus de la plaza*, Clasificación*, Integración*, Jornada*, Horas laboradas*, Turno opcional (SI/NO), Zona económica*, Servicio*, Antigüedad en el ISSSTE*, Antigüedad consecutiva*, Antigüedad en el ISSSTE con plaza de base*, Antigüedad federal*, Riesgos profesionales (SI/NO), Género*, Hijos (SI/NO), Escolaridad (hijos)*, Sindicalizado (SI/NO), Monto quincenal, Unidad del Monto, Conceptos sobre los que aplica (todas las claves en una misma celda, separados por comas), Monto mínimo, Unidad del monto mínimo, Monto máximo, Unidad del monto máximo, Fecha de inicio vigencia, Fecha de fin vigencia"
										Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Para este concepto se requiere: </B>" & sRequiredFields & "." & "</BR></BR><B>Notas:</B>&nbsp;a) Cargue los datos de preferencia en el orden especificado, b) En los campos marcados con * si su valor no es determinante, indiquelo con el valor 'Todos' o 'Todas', c) Para los campos que con (SI/NO) indique el valor 'SI' o 'NO'.</FONT>"
										Response.Write "<BR />"
										Response.Write "<BR />"
										Response.Write "<TEXTAREA NAME=""RawData"" ID=""RawDataTxtArea"" ROWS=""10"" COLS=""119"" CLASS=""TextFields"" onChange=""bReady = (this.value != '')""></TEXTAREA><BR /><BR />"
										Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"">"
									Response.Write "</FORM>"
									Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=" & oRequest("Action").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME>"
									Response.Write "<BR />"
									Response.Write "</DIV>"
									Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros que estan en proceso para el tipo de tabulador</FONT></A><BR /><BR />"
									Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
										Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
											Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
												Response.Write "<TR>"
													If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<TR><TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/></TD></TR>"
												Response.Write "</TR>"
											Response.Write "</TABLE><BR />"
											aConceptComponent(N_STATUS_ID_CONCEPT) = 0
											aConceptComponent(S_QUERY_CONDITION_CONCEPT) = ""
											lErrorNumber = DisplayConceptValuesTable(oRequest, oADODBConnection, iSelectedTab, False, sErrorDescription)
											If lErrorNumber <> 0 Then
												Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
												lErrorNumber = 0
												sErrorDescription = ""
											End If
										Response.Write "</FORM>"
									Response.Write "</DIV>"
								Response.Write "</TD>"
							End If
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						Case "EmployeeTypes"
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								lErrorNumber = DisplayEmployeeTypesTable(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
							Response.Write "</TD>"
						Case Else
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">" & vbNewLine
								If StrComp(oRequest("Action").Item, "UpdatePayroll", vbBinaryCompare) = 0 Then
									aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = " And (IsClosed<>1)"
									lErrorNumber = DisplayPayrollsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aPayrollComponent, sErrorDescription)
								Else
									lErrorNumber = DisplayPayrollsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aPayrollComponent, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
									Response.Write "<BR />"
								Else
									Response.Write "<BR />"
									If B_ISSSTE Then
										Select Case sAction
											Case "ModifyPayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" />"
											Case "ClosePayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" />"
											Case "DeletePayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" />"
												Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""360"" HEIGHT=""1"" />"
												Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Remove"" ID=""RemoveBtn"" VALUE=""Borrar Nómina"" CLASS=""RedButtons"" />"
										End Select
									Else
										Select Case sAction
											Case "ModifyPayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payroll.asp'"" />"
											Case "ClosePayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payroll.asp'"" />"
											Case "DeletePayroll"
												Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payroll.asp'"" />"
												Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""360"" HEIGHT=""1"" />"
												Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Remove"" ID=""RemoveBtn"" VALUE=""Borrar Nómina"" CLASS=""RedButtons"" />"
										End Select
									End If
								End If
							Response.Write "</TD>" & vbNewLine
					End Select
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
			End If
				Select Case sAction
					Case "AddPayroll"
						aPayrollComponent(N_TYPE_ID_PAYROLL) = 1
						lErrorNumber = DisplayPayrollForm(oRequest, oADODBConnection, GetASPFileName(""), aPayrollComponent, sErrorDescription)
					Case "AddSpecialPayroll"
						aPayrollComponent(N_TYPE_ID_PAYROLL) = -1
						lErrorNumber = DisplayPayrollForm(oRequest, oADODBConnection, GetASPFileName(""), aPayrollComponent, sErrorDescription)
					Case "ClosePayroll"
						If aPayrollComponent(N_ID_PAYROLL) = -1 Then
							Call DisplayModifyPayrollMessage(1, aPayrollComponent(N_ID_PAYROLL))
						Else
							If B_ISSSTE Then
								Call DisplayErrorMessage("Confirmación", "<FORM>La nómina del " & DisplayDateFromSerialNumber(aPayrollComponent(N_ID_PAYROLL), -1, -1, -1) & " fue cerrada.<BR /><INPUT TYPE=""BUTTON"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" /></FORM>")
							Else
								Call DisplayErrorMessage("Confirmación", "<FORM>La nómina del " & DisplayDateFromSerialNumber(aPayrollComponent(N_ID_PAYROLL), -1, -1, -1) & " fue cerrada.<BR /><INPUT TYPE=""BUTTON"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payroll.asp'"" /></FORM>")
							End If
						End If
					Case "Concepts"
						If iGlobalSectionID = 4 Then lErrorNumber = DisplayConceptForm(oRequest, oADODBConnection, GetASPFileName(""), aConceptComponent, sErrorDescription)
						'If lErrorNumber <> 0 Then
						'	Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						'	lErrorNumber = 0
						'	Response.Write "<BR />"
						'End If
						If Not IsEmpty(lSuccess) Then
							If lSuccess = 1 Then
								Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
							Else
								Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
							End If
						End If
					Case "ConceptsFile"
						If lErrorNumber = 0 Then
							Call DisplayErrorMessage("Confirmación", "El archivo para calcular la nómina fue generado con éxito")
						End If
					Case "ConceptsValues"
						If iGlobalSectionID = 4 Then
							'lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), True, iSelectedTab, aConceptComponent, sErrorDescription)
							lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), True, iSelectedTab, aConceptComponent, sErrorDescription)
							If Not IsEmpty(lSuccess) Then
								If lSuccess = 1 Then
									If Len(oRequest("iSuccessCount").Item) > 0 Then
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente. Se agregaron " & CLng(oRequest("iSuccessCount").Item) & " nuevos registros.")
									Else
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
									End If
								Else
									If Len(oRequest("iSuccessCount").Item) > 0 Then
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item) & "<BR> Solamente se registraron " & CLng(oRequest("iSuccessCount").Item) & " nuevos registros.")
									Else
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
									End If
								End If
							End If
						ElseIf Len(oRequest("Perceptions").Item) > 0 Then
							If lRecordID > 0 Then
								lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), False, iSelectedTab, aConceptComponent, sErrorDescription)
								If Not IsEmpty(lSuccess) Then
									If lSuccess = 1 Then
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
									Else
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
									End If
								End If
							End If
						Else
							lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), True, iSelectedTab, aConceptComponent, sErrorDescription)
							If Not IsEmpty(lSuccess) Then
								If lSuccess = 1 Then
									If Len(oRequest("iSuccessCount").Item) > 0 Then
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente. Se registraron " & CLng(oRequest("iSuccessCount").Item) & " nuevos registros.")
									Else
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
									End If
								Else
									If Len(oRequest("iSuccessCount").Item) > 0 Then
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item) & "<BR> Solamente se registraron " & CLng(oRequest("iSuccessCount").Item) & " nuevos registros.")
									Else
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
									End If
								End If
							End If
						End If
					Case "DeletePayroll"
						If aPayrollComponent(N_ID_PAYROLL) = -1 Then
							Response.Write "<IMG SRC=""Images/IcnInformationSmall.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
							Response.Write "<B>Seleccione la nómina a borrar</B>"
						Else
							If B_ISSSTE Then
								Call DisplayErrorMessage("Confirmación", "<FORM>La nómina del " & DisplayDateFromSerialNumber(aPayrollComponent(N_ID_PAYROLL), -1, -1, -1) & " fue cerrada.<BR /><INPUT TYPE=""BUTTON"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" /></FORM>")
							Else
								Call DisplayErrorMessage("Confirmación", "<FORM>La nómina del " & DisplayDateFromSerialNumber(aPayrollComponent(N_ID_PAYROLL), -1, -1, -1) & " fue cerrada.<BR /><INPUT TYPE=""BUTTON"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payroll.asp'"" /></FORM>")
							End If
						End If
					Case "EmployeeTypes"
					Case "ModifyPayroll"
						If Len(oRequest("CalculatePayroll").Item) > 0 Then
							sNames = L_NO_INSTRUCTIONS_FLAGS & "," & L_OPEN_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS
							lErrorNumber = DisplayFilterInformation(oRequest, sNames, False, "", sErrorDescription)
						End If
						Call DisplayModifyPayrollMessage(0, aPayrollComponent(N_ID_PAYROLL))
					Case "RetroactivePayroll"
					Case "UpdatePayroll"
						If aPayrollComponent(N_ID_PAYROLL) > -1 Then
							lErrorNumber = DisplayPayrollForm(oRequest, oADODBConnection, GetASPFileName(""), aPayrollComponent, sErrorDescription)
						End If
					Case Else
						Call GetNameFromTable(oADODBConnection, "Payrolls", lPayrollID, "", "", sNames, sErrorDescription)
						Response.Write "<BR /><TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							aMenuComponent(A_ELEMENTS_MENU) = Array(_
								Array("Nómina nueva",_
									  "Calcular una nueva nómina. Se calcularán los montos a pagar a partir de los conceptos de pago definidos por puestos, niveles, zonas económicas, etc. Posteriormente usted podrá agregar otros conceptos de pago para los empleados antes de cerrar la nómina.",_
									  "Images/MnPayroll.gif", "Payroll.asp?Action=AddPayroll", True),_
								Array("Generar prenómina",_
									  "Prepare todos los conceptos de pago para la revisión previa al cierre de la nómina. Esto le permitirá hacer cualquier corrección necesaria sobre la nómina '" & sNames & "'.",_
									  "Images/MnReportList.gif", "Payroll.asp?Action=ModifyPayroll", True),_
								Array("Cerrar nómina",_
									  "Si ya no se agregarán más conceptos de pago para los empleados, es necesario cerrar la nómina '" & sNames & "' para que se pueda pagar.",_
									  "Images/MnReports.gif", "Payroll.asp?Action=ClosePayroll", True),_
								Array("Pagos retroactivos",_
									  "Registrar ajustes a los conceptos de pago y calcular los montos retroactivos a partir de la fecha especificada.",_
									  "Images/MnPayments.gif", "Payroll.asp?Action=RetroactivePayroll", True),_
								Array("<LINE />",_
									  "",_
									  "", "", True),_
								Array("Conceptos de pago",_
									  "Administrar los valores de los conceptos de pago y por tipos de tabulador.",_
									  "Images/MnBudget.gif", "Payroll.asp?Action=Concepts", True)_
							)
							aMenuComponent(B_USE_DIV_MENU) = True
							Call DisplayMenuInTwoColumns(aMenuComponent)
						Response.Write "</TABLE><BR />"
				End Select
			If bShowTable Then
					Response.Write "</FONT></TD>" & vbNewLine
				Response.Write "</TR></TABLE>" & vbNewLine
			End If
		Else
			Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 1. </B>Introduzca el archivo a utilizar.<BR /><BR />"
			Select Case sAction
				Case "ConceptsValues"
					Select Case iStep
						Case 2
							Call DisplayConceptsValuesColumns(sFileName, lEmployeeTypeID, True, sErrorDescription)
						Case 3
							Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
							lErrorNumber = UploadConceptsValuesFile(oADODBConnection, sFileName, True, sErrorDescription)
							Response.Write "<BR />"
							If lErrorNumber = 0 Then
								Call DisplayErrorMessage("Confirmación", "Las becas de los hijos de los empleados fueron registrados con éxito.")
							Else
								Call DisplayErrorMessage("Error al registrar las becas de los hijos de los empleados.", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
					End Select
			End Select
		End If

		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>