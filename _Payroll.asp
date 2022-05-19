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
<!-- #include file="Libraries/_PayrollComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
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

iSelectedTab = -1
If Len(oRequest("Tab").Item) > 0 Then
	iSelectedTab = CInt(oRequest("Tab").Item)
ElseIf Len(oRequest("EmployeeTypeID").Item) > 0 Then
	iSelectedTab = CInt(oRequest("EmployeeTypeID").Item)
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
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			lErrorNumber = ModifyConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			lErrorNumber = RemoveConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			Redim aConceptComponent(N_CONCEPT_COMPONENT_SIZE)
			aConceptComponent(N_ID_CONCEPT) = -1
		ElseIf Len(oRequest("SetActive").Item) > 0 Then
			lErrorNumber = SetActiveForConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			Redim aConceptComponent(N_CONCEPT_COMPONENT_SIZE)
			aConceptComponent(N_ID_CONCEPT) = -1
			bShowForm = False
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
	Case "ConceptValues"
		If Len(oRequest("Add").Item) > 0 Then
			lErrorNumber = AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			Response.Redirect "Payroll.asp?Action=ConceptValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab 
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
					Response.Redirect "Payroll.asp?Action=ConceptValues&Success=1&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Perceptions=1&EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)
				Else
					Response.Redirect "Payroll.asp?Action=ConceptValues&Success=0&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Perceptions=1&EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)
				End If
			'End If
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			lErrorNumber = RemoveConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			aConceptComponent(N_RECORD_ID_CONCEPT) = -1
			Response.Redirect "Payroll.asp?Action=ConceptValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab
		ElseIf Len(oRequest("Authorize").Item) > 0 Then
			lErrorNumber = AuthorizeConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			aConceptComponent(N_RECORD_ID_CONCEPT) = -1
			Response.Redirect "Payroll.asp?Action=ConceptValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&EmployeeTypeID=" & iSelectedTab 
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
			If lErrorNumber = 0 Then
				lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 1, aPayrollComponent, sErrorDescription)
			End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Modificar nómina"
			lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
		ElseIf Len(oRequest("CalculatePayroll").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Prenómina"
			lErrorNumber = CalculatePayroll2(oRequest, oADODBConnection, 2, aPayrollComponent, sErrorDescription)
			If lErrorNumber =  0 Then
				lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				If lErrorNumber =  0 Then
					aPayrollComponent(N_CLOSED_PAYROLL) = 2
					lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				End If
			End If
		ElseIf Len(oRequest("DoClose").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Cerrar una nómina"
			lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 3, aPayrollComponent, sErrorDescription)
			If lErrorNumber =  0 Then
				lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
				If lErrorNumber =  0 Then
					aPayrollComponent(N_CLOSED_PAYROLL) = 1
					lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
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
				Case "Concepts", "ConceptsFile", "ConceptValues"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar un nuevo concepto",_
							  "",_
							  "", "Payroll.asp?Action=Concepts&New=1", (StrComp(sAction, "ConceptValues", vbBinaryCompare) <> 0)),_
						Array("Agregar un nuevo monto",_
							  "",_
							  "", "Payroll.asp?Action=ConceptValues&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptValues", vbBinaryCompare) = 0)),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptValues", vbBinaryCompare) = 0)),_
						Array("Exportar histórico a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & iGlobalSectionID & "&HistoryList=1&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aConceptComponent(N_ID_CONCEPT) > -1) And (StrComp(sAction, "ConceptValues", vbBinaryCompare) = 0)),_
						Array("Generar archivos de cálculo",_
							  "",_
							  "", "Payroll.asp?Action=ConceptsFile", True)_
					)
				'Case "ConceptValues"
				'	aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				'		Array("Agregar un nuevo tabulador",_
				'			  "",_
				'			  "", "Payroll.asp?Action=ConceptValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) <> -1)),_
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
					  "", "Payroll.asp?Action=ConceptValues&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Tab=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & "&New=1", (aConceptComponent(N_ID_CONCEPT) <> -1)),_
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
			Case "ConceptValues"
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
			Case Else
				If B_ISSSTE Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Nómina</B>"
				Else
					Response.Write "<B>Nómina</B>"
				End If
		End Select
		Response.Write "<BR /><BR />"

		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		End If
		If StrComp(sAction, "ConceptValues", vbBinaryCompare) = 0 Then
			If (Len(oRequest("Perceptions").Item) > 0) Or (iGlobalSectionID = 4) Then
				lErrorNumber = DisplayConceptValuesTabs(oRequest, aConceptComponent(N_ID_CONCEPT), iSelectedTab, sErrorDescription)
			End If
		End If
		If bShowTable Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
				Select Case sAction
					Case "Concepts", "ConceptsFile"
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
							If iGlobalSectionID <> 4 Then aConceptComponent(S_QUERY_CONDITION_CONCEPT) = " And (Concepts.ConceptID In (1,3,14,38,39,49,89))"
							lErrorNumber = DisplayConceptsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (iGlobalSectionID = 4), aConceptComponent, sErrorDescription)
						Response.Write "</TD>"
					Case "ConceptValues"
						If iGlobalSectionID = 4 Then
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								lErrorNumber = DisplayConceptValuesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, -1, True, False, aConceptComponent, sErrorDescription)
							Response.Write "</TD>"
						ElseIf Len(lRecordID) = 0 Then
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								If iGlobalSectionID = 3 Then
									lErrorNumber = DisplayConceptValuesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, iSelectedTab, True, False, aConceptComponent, sErrorDescription)
								Else
									lErrorNumber = DisplayConceptValuesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, -1, True, False, aConceptComponent, sErrorDescription)
								End If
							Response.Write "</TD>"
						End If
					Case "EmployeeTypes"
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
							lErrorNumber = DisplayEmployeeTypesTable(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
						Response.Write "</TD>"
					Case Else
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">" & vbNewLine
							lErrorNumber = DisplayPayrollsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aPayrollComponent, sErrorDescription)
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
				Case "ConceptsFile"
					If lErrorNumber = 0 Then
						Call DisplayErrorMessage("Confirmación", "El archivo para calcular la nómina fue generado con éxito")
					End If
				Case "ConceptValues"
					If iGlobalSectionID = 4 Then
						lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), True, iSelectedTab, aConceptComponent, sErrorDescription)
					ElseIf Len(oRequest("Perceptions").Item) > 0 Then
						If lRecordID > 0 Then
							lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), False, iSelectedTab, aConceptComponent, sErrorDescription)
						End If
					Else
						lErrorNumber = DisplayConceptValuesForm(oRequest, oADODBConnection, GetASPFileName(""), True, iSelectedTab, aConceptComponent, sErrorDescription)
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