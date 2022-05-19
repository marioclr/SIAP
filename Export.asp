<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponentConstants.asp" -->
<!-- #include file="Libraries/GraphComponent.asp" -->

<!-- #include file="Libraries/AbsenceComponent.asp" -->
<!-- #include file="Libraries/AlimonyTypeComponent.asp" -->
<!-- #include file="Libraries/AreaComponent.asp" -->
<!-- #include file="Libraries/BudgetComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/EmployeesLib.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/EmployeeSupportLib.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/Main_ISSSTELib.asp" -->
<!-- #include file="Libraries/PaymentComponent.asp" -->
<!-- #include file="Libraries/PaymentsLib.asp" -->
<!-- #include file="Libraries/PayrollComponent.asp" -->
<!-- #include file="Libraries/PayrollRevisionComponent.asp" -->
<!-- #include file="Libraries/PositionsLib.asp" -->
<!-- #include file="Libraries/PositionComponent.asp"-->
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/ReportComponent.asp" -->
<!-- #include file="Libraries/SADELibrary.asp" -->
<%
Dim sAction
Dim iSelectedTab
Dim bPrint
Dim bDummy
Dim sNames
Dim lReasonID
Dim sCondition
Dim iStatus
Dim iEmployeeTypeID
Dim sFilter

sFilter = ""
lReasonID = 0
If Len(oRequest("ReasonID").Item) > 0 Then lReasonID = CLng(oRequest("ReasonID").Item)
If Len(oRequest("EmployeeTypeID").Item)>0 Then
	iEmployeeTypeID = CLng(oRequest("EmployeeTypeID").Item)
Else
	iEmployeeTypeID = -1
End If
sAction = oRequest("Action").Item
bPrint = (Len(oRequest("Print").Item) > 0)
bDummy = (Len(oRequest("Dummy").Item) > 0)
Response.Cookies("SIAP_SectionID") = CInt(oRequest("SIAP_SectionID").Item)

If (aLoginComponent(N_USER_ID_LOGIN) > -1) And (Not bDummy) Then
	aOptionsComponent(L_ID_USER_OPTIONS) = aLoginComponent(N_USER_ID_LOGIN)
	Call GetOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
End If

Select Case sAction
	Case "Absences"
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	Case "Areas"
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	Case "BlockPayments"
	Case "Budgets"
	Case "CancelPayments"
	Case "ConceptsValues"
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	Case "Employees"
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
		Call GetEmployeesURLValues(oRequest, "", "", aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))
	Case "EmployeesForTaxAdjustment"
	Case "Jobs"
		Call InitializeJobComponent(oRequest, aJobComponent)
	Case "MedicalAreas"
	Case "ModifiedMoneys"
	Case "Positions"
			Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
			Call DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, "", "", "", sErrorDescription)
	Case "Payments"
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
		Call GetPaymentsURLValues(oRequest, "", aPaymentComponent(S_QUERY_CONDITION_PAYMENT))
	Case "Programs"
	Case "Reports"
		Call InitializeReportsComponent(oRequest, aReportsComponent)
	Case Else
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
		Call DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, "", "", "", sErrorDescription)
End Select
If lErrorNumber = 0 Then
	If Len(oRequest("Excel").Item) > 0 Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.CharSet = "iso-8859-1"
		Response.AddHeader "Content-Disposition", ("filename=Report" & GetSerialNumberForDate("") & ".xls;")
	ElseIf Len(oRequest("Word").Item) > 0 Then
		Response.ContentType = "application/vnd.ms-word"
		Response.CharSet = "iso-8859-1"
		Response.AddHeader "Content-Disposition", ("filename=Report" & GetSerialNumberForDate("") & ".doc;")
	ElseIf Len(oRequest("Text").Item) > 0 Then
		Response.ContentType = "text/plain"
		Response.CharSet = "iso-8859-1"
		Response.AddHeader "Content-Disposition", ("filename=Export.asp?" & oRequest & ";")
    ElseIf Len(oRequest("Web").Item) > 0 Then
		Response.ContentType = "text/plain"
		Response.CharSet = "iso-8859-1"
		Response.AddHeader "Content-Disposition", ("filename=Report" & GetSerialNumberForDate("") & ".html;")
	End If
End If
%>
<HTML>
	<HEAD>
		<%If (Len(oRequest("Excel").Item) = 0) And (Len(oRequest("Word").Item) = 0) Then%>
			<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/Export.css" />
			<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
			<!-- #include file="_JavaScript.asp" -->
		<%End If%>
		<TITLE><%
			Select Case sAction
				Case "Reports"
					Response.Write "Reporte: " & GetReportNameByConstant(aReportsComponent(N_ID_REPORTS))
			End Select
		%></TITLE>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If (Not bDummy) Then%><!-- #include file="_HeaderForExport.htm" --><%End If
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al exportar", sErrorDescription)
		End If
		If bPrint Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript"" SRC=""JavaScript/RollOver.js""></SCRIPT>"
			Response.Write "<DIV ID=""PrintDiv""></DIV>" & vbNewLine
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ExportHTMLToWindow(window.opener.document.all['" & oRequest("SourceContainer").Item & "'], true, true, window.document.all['PrintDiv']);" & vbNewLine
				If CInt(GetOption(aOptionsComponent, SHOW_PRINT_INFO_OPTION)) = 1 Then Response.Write "OpenNewWindow('PrintInstructions.asp', '', 'PrintInstructions', 500, 400, 'no', 'no');" & vbNewLine
				Response.Write "window.print();" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		Else
			Response.Write "<DIV ID=""ReportSection"">"
			Select Case sAction
				Case "EmployeeBeneficiaries"
					lErrorNumber = DisplayPendingEmployeesBeneficiariesTable(oRequest, oADODBConnection, True, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
				Case "Absences"
					lErrorNumber = DisplayPendingEmployeesAbscencesTable(oRequest, oADODBConnection, CInt(oRequest("Active").Item), True, lReasonID, sAction, aEmployeeComponent, sErrorDescription)
				Case "AlimonyTypes"
					lErrorNumber = DisplayAlimonyTypesTable(oRequest, oADODBConnection, True, sErrorDescription)
				Case "ApplyAbsences"
					lErrorNumber = DisplayAbsencesForApplyTable(oRequest, oADODBConnection, True, sErrorDescription)
				Case "Areas"
					If Len(oRequest("Excel").Item) > 0 Then
						If Len(oRequest("Complete").Item) > 0 Then
							lErrorNumber = DisplayAreasTableFull(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aAreaComponent, sErrorDescription)
						Else
							If Len(oRequest("PaymentCenters").Item) > 0 Then
								If Len(oRequest("ParentID").Item) > 0 Then
									Call GetAreaLevel(oRequest, oADODBConnection, CInt(oRequest("ParentID").Item), aAreaComponent, sErrorDescription)
									Call GetAreaParentID(oRequest, oADODBConnection, CInt(oRequest("ParentID").Item), aAreaComponent, sErrorDescription)
									aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & oRequest("ParentID").Item & ")"
								Else
									If Len(oRequest("ApplyFilter").Item) = 0 Then
										bPaymentCenters = True
										aAreaComponent(N_LEVEL_AREA) = 1
										If CInt(oRequest("ZoneID").Item) <> 6084 Then
											aAreaComponent(S_FILTER_CONDITION_AREA) = aAreaComponent(S_FILTER_CONDITION_AREA) & " And (Zones01.ZoneID=" & UCase(CStr(oRequest("ZoneID").Item)) & ")"
										End If
										If CInt(oRequest("GeneratingAreaID").Item) > 0 Then
											aAreaComponent(S_FILTER_CONDITION_AREA) = aAreaComponent(S_FILTER_CONDITION_AREA) & " And (Areas.GeneratingAreaID=" & UCase(CStr(oRequest("GeneratingAreaID").Item)) & ")"
										End If
										If CInt(oRequest("AreaCode").Item) > 0 Then
											aAreaComponent(S_FILTER_CONDITION_AREA) = aAreaComponent(S_FILTER_CONDITION_AREA) & " And (Areas.AreaCode = '" & UCase(CStr(oRequest("AreaCode").Item)) & "')"
										End If
										If CInt(oRequest("AreaShortName").Item) > 0 Then
											aAreaComponent(S_FILTER_CONDITION_AREA) = aAreaComponent(S_FILTER_CONDITION_AREA) & " And (Areas.AreaShortName = '" & UCase(CStr(oRequest("AreaShortName").Item)) & "')"
										End If
										If CInt(oRequest("CenterTypeID").Item) > 0 Then
											aAreaComponent(S_FILTER_CONDITION_AREA) = aAreaComponent(S_FILTER_CONDITION_AREA) & " And (Areas.CenterTypeID=" & CInt(oRequest("CenterTypeID").Item) & ")"
										End If
										aAreaComponent(S_QUERY_CONDITION_AREA) = aAreaComponent(S_QUERY_CONDITION_AREA) & aAreaComponent(S_FILTER_CONDITION_AREA)
									Else
										aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=-1)"
									End If
									'aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.PaymentCenterID>-1)"
								End If
							Else
								aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & oRequest("ParentID").Item & ")"
							End If
							lErrorNumber = DisplayAreasTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aAreaComponent, sErrorDescription)
						End If
					Else
						lErrorNumber = DisplayArea(oRequest, oADODBConnection, True, aAreaComponent, sErrorDescription)
					End If
				Case EMPLOYEES_BANK_ACCOUNTS
					lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, True, "EmployeesMovements", lReasonID, 1, aEmployeeComponent, sErrorDescription)
				Case "BlockPayments"
					lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 1, True, sErrorDescription)
				Case "Budgets"
					lErrorNumber = DisplayFullBudgetTable(oRequest, oADODBConnection, False, True, sErrorDescription)
				Case "CancelPayments"
					lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 0, True, sErrorDescription)
				Case "ConceptsValues"
					aConceptComponent(N_STATUS_ID_CONCEPT) = CInt(oRequest("Active").Item)
					lErrorNumber = DisplayConceptValuesTableSP(oRequest, oADODBConnection, iEmployeeTypeID, True, sErrorDescription)
				Case "Employees"
					If Len(oRequest("Excel").Item) > 0 Then
						If Len(oRequest("DoSearch").Item) = 0 Then
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
									Response.Write "<B>Número de empleado: </B>" & CleanStringForHTML(aEmployeeComponent(S_NUMBER_EMPLOYEE)) & "<BR />"
									Response.Write "<B>Nombre: </B>" & CleanStringForHTML(aEmployeeComponent(S_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & "<BR />"
									Response.Write "<B>RFC: </B>" & CleanStringForHTML(aEmployeeComponent(S_RFC_EMPLOYEE)) & "<BR />"
									Response.Write "<B>Plaza: </B>" & CleanStringForHTML(Right(("000000" & aEmployeeComponent(N_JOB_ID_EMPLOYEE)), Len("000000"))) & "<BR />"
									Call GetNameFromTable(oADODBConnection, "Positions", aEmployeeComponent(N_POSITION_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									Response.Write "<B>Puesto: </B>" & CleanStringForHTML(sNames) & "<BR />"
									Call GetNameFromTable(oADODBConnection, "Areas", aEmployeeComponent(N_AREA_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									Response.Write "<B>Adscripción: </B>" & CleanStringForHTML(sNames) & "<BR />"
									Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									Response.Write "<B>Estatus del empleado: </B>" & CleanStringForHTML(sNames) & "<BR />"
									Response.Write "<BR />"
								Response.Write "</FONT>"
							End If
						End If
						Select Case oRequest("Tab").Item
							Case 3
								lErrorNumber = DisplayEmployeeConceptsTable(oRequest, oADODBConnection, False, True, aEmployeeComponent, sErrorDescription)
							Case 4
								aAbsenceComponent(N_ACTIVE_ABSENCE) = 1
								lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aAbsenceComponent, sErrorDescription)
							Case 6
								Select Case CInt(oRequest("ReportID").Item)
									Case EMPLOYEE_HISTORY_LIST_REPORTS
										lErrorNumber = DisplayEmployeeHistoryList(oRequest, oADODBConnection, True, True, aEmployeeComponent, sErrorDescription)
									Case EMPLOYEE_FORM_HISTORY_LIST_REPORTS
										lErrorNumber = DisplayEmployeeFormHistoryList(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
									Case EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS
										lErrorNumber = DisplayEmployeePaymentsHistoryList(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
									Case EMPLOYEE_PAYROLL_REPORTS
										lErrorNumber = DisplayEmployeePayroll(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
									Case ISSSTE_1111_REPORTS
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
										lErrorNumber = DisplayJobHistoryList(oRequest, oADODBConnection, True, aJobComponent, sErrorDescription)
									Case JOBS_LIST_REPORTS 
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
										lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, True, aJobComponent, sErrorDescription)
								End Select
							Case 7
							Case Else
								lErrorNumber = DisplayEmployeesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aEmployeeComponent, sErrorDescription)
						End Select
					Else
						Select Case oRequest("Tab").Item
							Case 2
								If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
									lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									If lErrorNumber = 0 Then
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
										lErrorNumber = DisplayJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									End If
								End If
							Case 5
								If aEmployeeComponent(N_ID_EMPLOYEE) < 1000000 Then
									If aEmployeeComponent(N_ID_EMPLOYEE) >= 600000 Then
										lErrorNumber = BuildReport1110(oRequest, True, oADODBConnection, sErrorDescription)
									Else
										lErrorNumber = BuildReport1109(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
										'lErrorNumber = DisplayFormForEmployee(oADODBConnection, True, aEmployeeComponent, sErrorDescription)
									End If
								Else
									If aEmployeeComponent(N_ID_EMPLOYEE) >= 1600000 Then
										lErrorNumber = BuildReport1110(oRequest, True, oADODBConnection, sErrorDescription)
									Else
										lErrorNumber = BuildReport1109(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
										'lErrorNumber = DisplayFormForEmployee(oADODBConnection, True, aEmployeeComponent, sErrorDescription)
									End If
								End If
							Case Else
                                If Len(oRequest("ReportRH").Item)>0 Then
                                    lErrorNumber = DisplayEmployeeExport(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
                                Else
								    lErrorNumber = DisplayEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
                                End If
						End Select
					End If
				Case "EmployeesForTaxAdjustment"
					lErrorNumber = DisplayEmployeeTaxAdjustmentTable(oRequest, oADODBConnection, True, sErrorDescription)
				Case "EmployeesKardex"
					lErrorNumber = Display352SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
				Case "EmployeesKardex3"
					lErrorNumber = Display353SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
				Case "EmployeesKardex2"
					lErrorNumber = Display356SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
				Case "EmployeesMovements"
					If (lReasonID = EMPLOYEES_BANK_ACCOUNTS) Then
						lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, True, "EmployeesMovements", lReasonID, 1, aEmployeeComponent, sErrorDescription)
					ElseIf lReasonID = EMPLOYEES_GRADE Then
						lErrorNumber = DisplayEmployeesGradesTable(oRequest, oADODBConnection, True, aEmployeeComponent, sErrorDescription)
					ElseIf (lReasonID = EMPLOYEES_EXTRAHOURS) Or (lReasonID = EMPLOYEES_SUNDAYS) Then
						Select Case lReasonID
							Case EMPLOYEES_EXTRAHOURS
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 201
							Case EMPLOYEES_SUNDAYS
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 202
						End Select
						aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
						lErrorNumber = DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, True, "EmployeesMovements", lReasonID, aEmployeeComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, True, "EmployeesMovements", lReasonID, 0, aEmployeeComponent, sErrorDescription)
					End If
				Case "Jobs"
					aJobComponent(S_QUERY_CONDITION_JOB) = " And (Jobs.JobID=" & oRequest("JobNumber").Item & ")"
					aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_AREA
					lErrorNumber = DisplayJobsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aJobComponent, sErrorDescription)
				Case "MedicalAreas"
					lErrorNumber = DisplayMedicalAreasTable(oRequest, oADODBConnection, True, 1, aEmployeeComponent, sErrorDescription)
				Case "ModifiedMoneys"
					lErrorNumber = PrintModifiedMoneys(oRequest, oADODBConnection, sErrorDescription)
				Case "PaperworkList"
					lErrorNumber = PrintPaperworkList(oRequest, oADODBConnection, oRequest("ListID").Item, sErrorDescription)
				Case "Paperworks"
					Call GetPaperworksURLValues(oRequest, False, True, sCondition)
					lErrorNumber = DisplayPaperworksForSupportTable(oRequest, oADODBConnection, False, True, sCondition, sErrorDescription)
				Case "Payments"
					If Len(oRequest("Excel").Item) > 0 Then
						lErrorNumber = DisplayPaymentsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aPaymentComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
					End If
				Case "PaymentsRecords", "Replacement"
					aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords"
					Call DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, False, oRequest("Action").Item, "", sErrorDescription)
					lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, True, aCatalogComponent, sErrorDescription)
				Case "PayrollRevision"
					lErrorNumber = DisplayPayrollRevisionTable(oRequest, oADODBConnection, True, aPayrollRevisionComponent, sErrorDescription)
				Case "Positions"
					If 	Len(oRequest("Excel").Item) > 0  Then
						If Len(oRequest("PositionID").Item) > 0 Then
							lErrorNumber = DisplayPosition(oADODBConnection, oRequest("PositionID").Item,oRequest("StartDate"), True, sErrorDescription)
						Else
						    'lErrorNumber = DisplayTables(sAction, sErrorDescription)
                            lErrorNumber = DisplayPositionsTable(oRequest, oADODBConnection, True, aPositionComponent, sErrorDescription)
						End If
					Else
						lErrorNumber = DisplayPosition(oADODBConnection, oRequest("PositionID").Item,oRequest("StartDate"), True, sErrorDescription)
					End If
				Case "Programs"
					lErrorNumber = DisplayFullProgramTable(oRequest, oADODBConnection, oRequest("ProgramYear").Item, False, True, sErrorDescription)
				Case "Reexpedition"
					aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords2"
					Call DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, False, oRequest("Action").Item, "", sErrorDescription)
					lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, True, aCatalogComponent, sErrorDescription)
				Case "Reports"
					Select Case aReportsComponent(N_ID_REPORTS)
						Case LOGS_HISTORY_REPORTS
							sFlags = L_NO_DIV_FLAGS & "," & L_LOG_DATE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayLogCount(oRequest, True, sErrorDescription)
								Response.Write "<BR />"
								lErrorNumber = DisplayLogHistoryList(oRequest, True, sErrorDescription)
							End If
						Case AREAS_COUNT_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_CENTER_SUBTYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ZONE_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS & "," & L_AREA_STATUS_FLAGS & "," & L_AREA_ACTIVE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayAreasCount(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case EMPLOYEES_COUNT_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayEmployeesCount(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case JOBS_COUNT_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayJobsCount(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case AREAS_LIST_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayAreasList(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case EMPLOYEES_LIST_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayEmployeesList(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case JOBS_LIST_REPORTS, SPECIAL_JOBS_LIST_REPORTS, JOBS_LIST_BY_MODIFY_DATE
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayJobsList(oRequest, oADODBConnection, True, sErrorDescription)
							End If

						Case ISSSTE_1001_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS & "," & L_REPORT_TITLE_FLAGS
							asTitles = Split("Título 1;;;Título 2;;;Título 3", LIST_SEPARATOR)
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
'								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1001(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1002_REPORTS
							sFlags = L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1002(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1004_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1004(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1007_REPORTS
							sFlags = L_DONT_CLOSE_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATE_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1007(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1009_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1009(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1012_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1012(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1013_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1013(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1014_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1014(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1015_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1015(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1018_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1018(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1019_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1019(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1020_REPORTS
							sFlags = L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1020(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1021_REPORTS
							sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1021(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1102_REPORTS
							sFlags = L_DATE_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_AREA_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1102(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1116_REPORTS, ISSSTE_1204_REPORTS, ISSSTE_1702_REPORTS
							sFlags = L_EMPLOYEE_NUMBER1_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								'lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1116(oRequest, oADODBConnection, True, Null, sErrorDescription)
							End If
						Case ISSSTE_1117_REPORTS, ISSSTE_1205_REPORTS, ISSSTE_1703_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1117(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1118_REPORTS, ISSSTE_1206_REPORTS, ISSSTE_1704_REPORTS
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							Response.Write "<B>Fecha de cálculo: " & DisplayDateFromSerialNumber(oRequest("PayrollYear").Item & oRequest("PayrollMonth").Item & oRequest("PayrollDay").Item, -1, -1, -1) & "</B>"
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1118(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1119_REPORTS
							sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							Response.Write "<B>Fecha de revisión: " & DisplayDateFromSerialNumber(oRequest("PayrollYear").Item & oRequest("PayrollMonth").Item & oRequest("PayrollDay").Item, -1, -1, -1) & "</B>"
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1119(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1207_REPORTS
							sFlags = L_EMPLOYEE_NUMBER1_FLAGS
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1207(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1208_REPORTS
							sFlags = L_EMPLOYEE_NUMBER1_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_DATE_FLAGS
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1208(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1335_REPORTS
							sFlags = L_EMPLOYEE_TYPE1_FLAGS & "," & L_CONCEPTS_VALUES_STATUS_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1335(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1336_REPORTS
							sFlags = L_GENERATING_AREAS_FLAGS & "," & L_AREA_CODE_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_CENTER_SUBTYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1336(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1337_REPORTS
							sFlags = L_GENERATING_AREAS_FLAGS & "," & L_AREA_CODE_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1337(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1354_REPORTS
							lErrorNumber = BuildReport1354(oRequest, oADODBConnection, True, sErrorDescription)
						Case ISSSTE_1356_REPORTS
							lErrorNumber = BuildReport1356(oRequest, oADODBConnection, True, sErrorDescription)
						Case ISSSTE_1364_REPORTS
							sFlags = L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_COURSE_NAME_FLAGS & "," & L_COURSE_DIPLOMA_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1364(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1365_REPORTS
							If lErrorNumber = 0 Then
								Call GetNameFromTable(oADODBConnection, SADE_PREFIX & "Curso", CLng(oRequest("CourseID").Item), "", "", sNames, sErrorDescription)
								Response.Write "Calificaciones obtenidas por los empleados para el curso <B>""" & sNames & """</B>.<BR /><BR />"
								lErrorNumber = DisplayCourseEmployeesTable(oRequest, oSIAPSADEADODBConnection, CLng(oRequest("CourseID").Item), False, True, sErrorDescription)
							End If
						Case ISSSTE_1367_REPORTS
							If lErrorNumber = 0 Then
								lErrorNumber = DisplayEmployeeCurriculum(oRequest, oSIAPSADEADODBConnection, CLng(oRequest("EmployeeID").Item), True, sErrorDescription)
							End If
						Case ISSSTE_1369_REPORTS
							Call DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, True, True, False, sErrorDescription)
							lErrorNumber = Display369SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
						Case ISSSTE_1411_REPORTS
							sFlags = L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1411(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1417_REPORTS
							sFlags = L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1417(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1420_REPORTS, ISSSTE_2420_REPORTS
							sFlags = L_ONE_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1420(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1421_REPORTS, ISSSTE_2421_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1421(oRequest, oADODBConnection, False, True, sErrorDescription)
							End If
						Case ISSSTE_1422_REPORTS, ISSSTE_2422_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1422(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1423_REPORTS, ISSSTE_2423_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1421(oRequest, oADODBConnection, True, True, sErrorDescription)
							End If
						Case ISSSTE_1424_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1424(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1425_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1425(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1427_REPORTS, ISSSTE_2427_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1427(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1428_REPORTS, ISSSTE_2428_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1428(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1429_REPORTS, ISSSTE_2429_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1429(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1430_REPORTS, ISSSTE_2430_REPORTS
							sFlags = L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1430(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1471_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ONE_BANK_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1471(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1114_REPORTS, ISSSTE_1472_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
									Response.Write "<B>Tipo de pago:</B><BR />"
									Select Case oRequest("PaymentType").Item
										Case ""
											Response.Write "-Todos"
										Case "0"
											Response.Write "-Cheques"
										Case "1"
											Response.Write "-Depósitos"
									End Select
								Response.Write "</FONT><BR /><BR />"
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1472(oRequest, oADODBConnection, -1, True, sErrorDescription)
							End If
						Case ISSSTE_1473_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1472(oRequest, oADODBConnection, 1, True, sErrorDescription)
							End If
						Case ISSSTE_1490_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_BANK_FLAGS & "," & L_STATE_TYPE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
'								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1490(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1494_REPORTS
							sFlags = L_CLOSED_PAYROLL_FLAGS & "," & L_MEMORY_CONCEPT_ID_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
							'	lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1494(oRequest, oADODBConnection, oRequest("ConceptID").Item, True, sErrorDescription)
							End If
						Case ISSSTE_1499_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1499(oRequest, oADODBConnection, oRequest("EmployeeNumber").Item, sErrorDescription)
							End If
						Case ISSSTE_1503_REPORTS
							sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), True, sErrorDescription)
							End If
						Case ISSSTE_1504_REPORTS, ISSSTE_1701_REPORTS
							sFlags = L_BUDGET_AREA_FLAGS & "," & L_BUDGET_GROUP_DUTY_FLAGS & "," & L_BUDGET_FUND_FLAGS & "," & L_BUDGET_DUTY_FLAGS & "," & L_BUDGET_ACTIVE_DUTY_FLAGS & "," & L_BUDGET_SPECIFIC_DUTY_FLAGS & "," & L_BUDGET_PROGRAM_FLAGS & "," & L_BUDGET_REGION_FLAGS & "," & L_BUDGET_UR_FLAGS & "," & L_BUDGET_CT_FLAGS & "," & L_BUDGET_AUX_FLAGS & "," & L_BUDGET_LOCATION_FLAGS & "," & L_BUDGET_BUDGET1_FLAGS & "," & L_BUDGET_BUDGET2_FLAGS & "," & L_BUDGET_BUDGET3_FLAGS & "," & L_BUDGET_CONFINE_TYPE_FLAGS & "," & L_BUDGET_ACTIVITY1_FLAGS & "," & L_BUDGET_ACTIVITY2_FLAGS & "," & L_BUDGET_PROCESS_FLAGS & "," & L_BUDGET_YEAR_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1504(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1561_REPORTS
							sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), True, sErrorDescription)
							End If
						Case ISSSTE_1562_REPORTS
							sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), True, sErrorDescription)
							End If
						Case ISSSTE_1563_REPORTS
							sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), True, sErrorDescription)
							End If
						Case ISSSTE_1571_REPORTS
							sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), True, sErrorDescription)
							End If
						Case ISSSTE_1581_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1581(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1582_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1582(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1583_REPORTS
							sFlags = L_YEARS_FLAGS & "," & L_COMPANY_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1583(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1584_REPORTS
							sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1584(oRequest, oADODBConnection, True, sErrorDescription)
							End If
						Case ISSSTE_1600_REPORTS
							If lErrorNumber = 0 Then
								lErrorNumber = PrintPaperwork(oRequest, oADODBConnection, CLng(oRequest("PaperworkID").Item), -1, sErrorDescription)
							End If
						Case ISSSTE_1602_REPORTS
							If lErrorNumber = 0 Then
								lErrorNumber = PrintPaperworkGuide(oRequest, oADODBConnection, CLng(oRequest("PaperworkID").Item), CLng(oRequest("AddressID1").Item), CLng(oRequest("AddressID2").Item), sErrorDescription)
							End If
						Case ISSSTE_1613_REPORTS
							sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_OWNERS_FLAGS & "," & L_ZIP_WARNING_FLAGS
							If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, True, "", sErrorDescription)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = BuildReport1613(oRequest, oADODBConnection, True, sErrorDescription)
							End If
					End Select
				Case "Zones"
					If Len(oRequest("ParentID").Item) > 0 Then
						aZoneComponent(S_QUERY_CONDITION_ZONE) = " And (ParentID=" & oRequest("ParentID").Item & ")"
					Else
						aZoneComponent(S_QUERY_CONDITION_ZONE) = " And (ParentID=-1)"
					End If
					lErrorNumber = DisplayZonesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, True, aZoneComponent, sErrorDescription)
				Case Else
					If InStr(1, ",267,423,424,425,426,", "," & oRequest("SectionID").Item & ",", vbBinaryCompare) > 0 Then
						Call DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, True, True, False, sErrorDescription)
						If StrComp(oRequest("Internal").Item, "1", vbBinaryCompare) = 0 Then
							Select Case CInt(oRequest("SectionID").Item)
								Case 423
									Response.Write "<B>Registros de guardias para los empleados internos</B><BR /><BR />"
								Case 424
									Response.Write "<B>Registros de suplencias para los empleados internos</B><BR /><BR />"
								Case 425
									Response.Write "<B>Registros de rezago quirúrgico para los empleados internos</B><BR /><BR />"
								Case 426
									Response.Write "<B>Registros de PROVAC para los empleados internos</B><BR /><BR />"
							End Select
						Else
							Select Case CInt(oRequest("SectionID").Item)
								Case 423
									Response.Write "<B>Registros de guardias para los empleados externos</B><BR /><BR />"
								Case 424
									Response.Write "<B>Registros de suplencias para los empleados externos</B><BR /><BR />"
								Case 425
									Response.Write "<B>Registros de rezago quirúrgico para los empleados externos</B><BR /><BR />"
								Case 426
									Response.Write "<B>Registros de PROVAC para los empleados externos</B><BR /><BR />"
							End Select
						End If
					End If
					lErrorNumber = DisplayTables(sAction, sErrorDescription)
			End Select
			Response.Write "</DIV>"
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessageInPlainText("Mensaje del sistema", sErrorDescription, "<BR />")
			End If
		End If
		If (Not bDummy) Then%><!-- #include file="_FooterForExport.htm" --><%End If%>
	</BODY>
</HTML>