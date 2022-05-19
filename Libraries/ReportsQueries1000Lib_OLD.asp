<%
Function BuildReport1001(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Hoja de cifras. Se genera a partir de que se crea una nómina nueva
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1001"
	Dim sContents
	Dim sTitle
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sBoldBegin
	Dim sBoldEnd
	Dim oRecordset
	Dim iCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "PaymentCenters.AreaID", "EmployeesHistoryList.PaymentCenterID")
	sErrorDescription = "No se pudieron obtener los totales por conceptos de pago para la nómina especificada."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, IsDeduction, OrderInList, ConceptShortName, ConceptName, Sum(ConceptAmount) As TotalForConcept From BankAccounts, Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, IsDeduction, OrderInList, ConceptShortName, ConceptName Order By IsDeduction, OrderInList, ConceptShortName, ConceptName", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1007.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
			sContents = Replace(sContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
			Response.Write sContents
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><CENTER><B>" & asTitles(CInt(oRequest("ReportTitle").Item)) & "<BR /><BR /></B></CENTER></FONT>"

			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Concepto,Monto", ",", -1, vbBinaryCompare)
				asCellWidths = Split("400,400", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",RIGHT", ",", -1, vbBinaryCompare)
				iCounter = 1
				Do While Not oRecordset.EOF
					If (iCounter = 40) And (Len(oRequest("Word").Item) > 0) Then
						Response.Write "</TABLE>"
						Response.Write sContents
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><CENTER><B>" & asTitles(CInt(oRequest("ReportTitle").Item)) & "<BR /><BR /></B></CENTER></FONT>"

						Response.Write "<TABLE BORDER="""
							If Not bForExport Then
								Response.Write "0"
							Else
								Response.Write "1"
							End If
						Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
							asColumnsTitles = Split("Concepto,Monto", ",", -1, vbBinaryCompare)
							asCellWidths = Split("400,400", ",", -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
							Else
								If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
									lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								Else
									lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								End If
							End If

							asCellAlignments = Split(",RIGHT", ",", -1, vbBinaryCompare)
					End If
					If CInt(oRecordset.Fields("ConceptID").Value) > 0 Then
						sBoldBegin = ""
						sBoldEnd = ""
					Else
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("TotalForConcept").Value), 2, True, False, True) & sBoldEnd

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					iCounter = iCounter + 1
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1001 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1002(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Hoja de cifras. Se genera a partir de que se crea una nómina nueva
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1002"
	Dim lCurrentID
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPreviousPayrollID
	Dim sCondition
	Dim dDiff
	Dim sFontBegin
	Dim sFontEnd
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	sErrorDescription = "No se pudo obtener la nómina anterior a la nómina especificada."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(PayrollID) As PreviousPayrollID From Payrolls Where (PayrollID<" & lPayrollID & ")", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lPreviousPayrollID = CLng(oRecordset.Fields("PreviousPayrollID").Value)
			oRecordset.Close

			If Len(lPreviousPayrollID) > 0 Then
				sCondition = ""
				If Len(oRequest("EmployeeID").Item) Then sCondition = " And (Employees.EmployeeID=" & oRequest("EmployeeID").Item & ")"
				sErrorDescription = "No se pudieron obtener los totales por conceptos de pago para la nómina especificada."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, EmployeeName, EmployeeLastName, EmployeeLastName2, ConceptName, Payroll_" & lPayrollID & ".ConceptAmount, Payroll_" & lPreviousPayrollID & ".ConceptAmount As PreviousAmount From Payroll_" & lPayrollID & ", Payroll_" & lPreviousPayrollID & ", Employees, Concepts Where (Payroll_" & lPayrollID & ".EmployeeID=Payroll_" & lPreviousPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Payroll_" & lPreviousPayrollID & ".ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".ConceptID<=0) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeeLastName, EmployeeLastName2, EmployeeName, OrderInList", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TABLE WIDTH=""800"" BORDER="""
							If Not bForExport Then
								Response.Write "0"
							Else
								Response.Write "1"
							End If
						Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
							asColumnsTitles = Split("Empleado,Concepto,Monto,Monto anterior,Diferencia", ",", -1, vbBinaryCompare)
							asCellWidths = Split("200,150,150,150,150", ",", -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
							Else
								If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
									lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								Else
									lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								End If
							End If

							asCellAlignments = Split(",,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
							lCurrentID = -1
							Do While Not oRecordset.EOF
								sFontBegin = ""
								sFontEnd = ""
								dDiff = Abs(CDbl(oRecordset.Fields("ConceptAmount").Value) - CDbl(oRecordset.Fields("PreviousAmount").Value))
								If ((dDiff / CDbl(oRecordset.Fields("ConceptAmount").Value)) > 0.1) Or ((dDiff / CDbl(oRecordset.Fields("PreviousAmount").Value)) > 0.1) Then
									sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>"
									sFontEnd = "</B></FONT>"
								End If
								If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value)) & sFontEnd
									Else
										sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)) & sFontEnd
									End If
									lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								Else
									sRowContents = "&nbsp;"
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & FormatNumber(CDbl(oRecordset.Fields("PreviousAmount").Value), 2, True, False, True) & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & FormatNumber(dDiff, 2, True, False, True) & sFontEnd

								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
						Response.Write "</TABLE>"
					Else
						lErrorNumber = L_ERR_NO_RECORDS
						sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
					End If
				End If
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1002 = lErrorNumber
	Err.Clear
End Function

Function BuildReports1003(oRequest, oADODBConnection, bReview, sErrorDescription)
'************************************************************
'Purpose: Listado de firmas masivo. Reporte basado en la hoja 001157
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bReview
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReports1003"
	Const N_EMPLOYEES_PER_FILE = 500
	Dim sGeneralHeader
	Dim sEmployeeHeader
	Dim sEmployeeHeaderTotals
	Dim sContents
	Dim sCondition
	Dim sOrderBy
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim lStartPayrollDate
	Dim lIssueDate
	Dim lCurrentCompanyID
	Dim lCurrentPaymentCenterID
	Dim sCurrentEmployeeID
	Dim sCurrentID
	Dim sTempCurrent
	Dim asStateNames
	Dim asPath
	Dim sPeriod
	Dim asConceptsP
	Dim asConceptsD
	Dim sTags
	Dim sCodes
	Dim asMonths
	Dim dTemp
	Dim iIndex
	Dim jIndex
	Dim iLast
	Dim iBound
	Dim jBound
	Dim sRowContents
	Dim sConcepts
	Dim oRecordset
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim lCounter
	Dim adTotal
	Dim lURCounter
	Dim lURPageCount
	Dim adURTotal
	Dim lCTCounter
	Dim lCTPageCount
	Dim bCTFirst
	Dim adCTTotal
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	Dim lTempID
	Dim lTempDeduction
	Dim sTempShortName
	Dim dTempAmount
	Dim lTempStartDate
	Dim lTempEndDate

	Dim i
	Dim lPageCount
	Dim bPageHeader
	Dim lRowsCountPerPage
	Dim lRowsCountPerEmployee
	Dim sOriginalHeaderData
	Dim sSectionSummaryData
	Dim sHeaderDataForEmployee
	Dim lHeaderCount
	Dim sEmployeeData
	Dim lEmployeeConceptsCount
	Dim lPageBreak
	Dim bFirst
	Dim bSection
	Dim bSectionSummary
	Dim lEmployeeCount
	Dim iFileCount
	Dim sFolderPath
	Dim sZipFile

	Const N_ROWS_PER_PAGE = 57
	Const N_ROWS_PER_HEADER = 10
	Const N_ROWS_FOR_SUMMARY = 5
	Const N_ROWS_FOR_SECTION_SUMMARY = 2
	Const N_ROWS_FOR_EMPLOYEE_DATA = 7
	asMonths = Split(",", ",")
	asMonths(0) = Split("11,12,01,02,03,04,05,06,07,08,09,10,11", ",")

	asMonths(1) = Split("09,10,11,12,01,02,03,04,05,06,07,08,09", ",")

	If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		If bReview Then
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_69ar.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_69br.htm"), sErrorDescription)
			sEmployeeHeaderTotals = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_69brt.htm"), sErrorDescription)
		Else
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_69a.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_69b.htm"), sErrorDescription)
			sEmployeeHeaderTotals = sEmployeeHeader
		End If
	ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		If bReview Then
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_154ar.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_154br.htm"), sErrorDescription)
			sEmployeeHeaderTotals = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_154brt.htm"), sErrorDescription)
		Else
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_154a.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_154b.htm"), sErrorDescription)
			sEmployeeHeaderTotals = sEmployeeHeader
		End If
	ElseIf StrComp(oRequest("CheckConceptID").Item, "11", vbBinaryCompare) = 0 Then
		If bReview Then
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_11ar.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_11br.htm"), sErrorDescription)
			sEmployeeHeaderTotals = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_11brt.htm"), sErrorDescription)
		Else
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_11a.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_11b.htm"), sErrorDescription)
			sEmployeeHeaderTotals = sEmployeeHeader
		End If
	Else
		If bReview Then
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_1r.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_2r.htm"), sErrorDescription)
			sEmployeeHeaderTotals = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_2rt.htm"), sErrorDescription)
		Else
			sGeneralHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_1.htm"), sErrorDescription)
			sEmployeeHeader = GetFileContents(Server.MapPath("Templates\HeaderForReport_1003_2.htm"), sErrorDescription)
			sEmployeeHeaderTotals = sEmployeeHeader
		End If
	End If
	If (Len(sGeneralHeader) > 0) And (Len(sEmployeeHeader) > 0) Then
		oStartDate = Now()
		sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asStateNames = ""
			Do While Not oRecordset.EOF
				asStateNames = asStateNames & LIST_SEPARATOR & SizeText(CStr(CleanStringForHTML(oRecordset.Fields("ZoneName").Value)), " ", 19, 1)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asStateNames = Split(asStateNames, LIST_SEPARATOR)

			Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
			lPayrollID = CLng(lPayrollID)
			sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies", "EmployeesHistoryList"), "EmployeeTypes", "EmployeesHistoryList"), "Payroll_YYYYMMDD", "Payroll_" & lPayrollID), "(Zones.", "(AreasZones.")
			If Len(oRequest("EmployeeID").Item) > 0 Then sCondition = sCondition & " And (Employees.EmployeeID In (" & Replace(oRequest("EmployeeID").Item, " ", "") & "))"
			lStartPayrollDate = GetPayrollStartDate(lForPayrollID)
			lIssueDate = CLng(oRequest("PayrollIssueYear").Item & oRequest("PayrollIssueMonth").Item & oRequest("PayrollIssueDay").Item)
			If lIssueDate = 0 Then lIssueDate = lPayrollID
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'				sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
				sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
			End If

			Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

			If bReview Then
				sOrderBy = "CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, RecordDate, OrderInList, RecordID"
			Else
				sOrderBy = "CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, OrderInList, RecordDate, RecordID"
			End If
			sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
			If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
				sOrderBy = Replace(sOrderBy, "EmployeesHistoryList.EmployeeNumber", "EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber")
				If bPayrollIsClosed Then
					If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
						sOrderBy = "CheckNumber, " & sOrderBy
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, Payments, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, Payments, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
					Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy & " -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
				sOrderBy = Replace(sOrderBy, "EmployeesHistoryList.EmployeeNumber", "EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber")
				If bPayrollIsClosed Then
					If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
						sOrderBy = "CheckNumber, " & sOrderBy
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesHistoryListForPayroll, EmployeesCreditorsLKP, Employees, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Payroll_" & lPayrollID & ", Concepts, Payments, EmployeesChangesLKP As EmpChLKP, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payments.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lPayrollID & ".RecordID) And (EmployeesCreditorsLKP.EmployeeID = EmpChLKP.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Payroll_" & lPayrollID & ".ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (EmployeesHistoryListForPayroll.PayrollID = " & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.AreaID = Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (PaymentCenters.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate <= " & lPayrollID & ") And (EmployeesCreditorsLKP.EndDate >= " & lPayrollID & ") And (companies.StartDate <= " & lPayrollID & ") And (Companies.EndDate >= " & lPayrollID & ") And (Zones.StartDate<=" & lPayrollID & ") And (Zones.EndDate>=" & lPayrollID & ") And (PaymentCenters.StartDate <= " & lPayrollID & ") And (PaymentCenters.EndDate >= " & lPayrollID & ") And (Areas.StartDate <= " & lPayrollID & ") And (Areas.EndDate >= " & lPayrollID & ") And (Positions.StartDate <= " & lPayrollID & ") And (Positions.EndDate >= " & lPayrollID & ") And (Levels.StartDate <= " & lPayrollID & ") And (Levels.EndDate >= " & lPayrollID & ") And (GroupGradeLevels.StartDate <= " & lPayrollID & ") And (GroupGradeLevels.EndDate >= " & lPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesHistoryListForPayroll, EmployeesCreditorsLKP, Employees, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Payroll_" & lPayrollID & ", Concepts, Payments, EmployeesChangesLKP As EmpChLKP, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payments.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lPayrollID & ".RecordID) And (EmployeesCreditorsLKP.EmployeeID = EmpChLKP.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Payroll_" & lPayrollID & ".ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (EmployeesHistoryListForPayroll.PayrollID = " & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.AreaID = Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (PaymentCenters.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate <= " & lPayrollID & ") And (EmployeesCreditorsLKP.EndDate >= " & lPayrollID & ") And (companies.StartDate <= " & lPayrollID & ") And (Companies.EndDate >= " & lPayrollID & ") And (Zones.StartDate<=" & lPayrollID & ") And (Zones.EndDate>=" & lPayrollID & ") And (PaymentCenters.StartDate <= " & lPayrollID & ") And (PaymentCenters.EndDate >= " & lPayrollID & ") And (Areas.StartDate <= " & lPayrollID & ") And (Areas.EndDate >= " & lPayrollID & ") And (Positions.StartDate <= " & lPayrollID & ") And (Positions.EndDate >= " & lPayrollID & ") And (Levels.StartDate <= " & lPayrollID & ") And (Levels.EndDate >= " & lPayrollID & ") And (GroupGradeLevels.StartDate <= " & lPayrollID & ") And (GroupGradeLevels.EndDate >= " & lPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
					Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.concepts40 From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP EmpChLKP, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Positions, Levels, GroupGradeLevels, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lForPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lForPayrollID & ".RecordID)  And (EmployeesHistoryListForPayroll.EmployeeID = EmpChLKP.EmployeeID) And (employeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesHistoryListForPayroll.AreaID = PaymentCenters.AreaID) And (PaymentCenters.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (Payroll_20120630.ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PayrollID = " & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate <= " & lForPayrollID & ")  And (EmployeesCreditorsLKP.EndDate >= " & lForPayrollID & ") And (companies.StartDate <= " & lForPayrollID & ") And (Companies.EndDate >= " & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate <= " & lForPayrollID & ") And (PaymentCenters.EndDate >= " & lForPayrollID & ") And (Areas.StartDate <= " & lForPayrollID & ") And (Areas.EndDate >= " & lForPayrollID & ") And (Positions.StartDate <= " & lForPayrollID & ") And (Positions.EndDate >= " & lForPayrollID & ") And (Levels.StartDate <= " & lForPayrollID & ") And (Levels.EndDate >= " & lForPayrollID & ") And (GroupGradeLevels.StartDate <= " & lForPayrollID & ") And (GroupGradeLevels.EndDate >= " & lForPayrollID & ") And (EmpChLKP.PayrollID = " & lForPayrollID & ") And (EmpChLKP.PayrollDate = Payroll_" & lForPayrollID & ".RecordDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"ParentAreas.AreaCode,",""), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesCreditorsLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesCreditorsLKP.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.CompanyID, '0' As EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesCreditorsLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesCreditorsLKP.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.CompanyID, '0' As EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesCreditorsLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesCreditorsLKP.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy & " -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
						If StrComp(oRequest("CheckConceptID").Item, "11", vbBinaryCompare) = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						End If
					Else
'						If oRecordset.EOF And (InStr(1, sCondition, "Payments.", vbBinaryCompare) = 0) Then
							If StrComp(oRequest("CheckConceptID").Item, "11", vbBinaryCompare) = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmpChLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
							End If
'						End If
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.CompanyID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.CompanyID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesChangesLKP As EmpChLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By " & sOrderBy & " -->" & vbNewLine
				End If
			End If
			iFileCount=0
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sDate = GetSerialNumberForDate("")
					sFolderPath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
					sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
					lErrorNumber = CreateFolder(Server.MapPath(sFolderPath), sErrorDescription)
					sZipFile = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
					If lErrorNumber = 0 Then
						'sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
						sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
						If lErrorNumber = 0 Then
							Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sZipFile) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
							Response.Flush()
							lCurrentCompanyID = -1
							lCurrentPaymentCenterID = -1
							sCurrentEmployeeID = "-1"
							sCurrentID = "-1"
							sContents = ""
							sConcepts = ""
							sTags = ""
							sCodes = ""
							lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<FONT FACE=""Courier"" SIZE=""3""><PRE>", sErrorDescription)
								' ****** Inicialización de variables generales para el reporte
								lHeaderCount = 0
								lRowsCountPerPage = 0
								lCounter = 0
								adTotal = Split("0,0,0", ",")
								adTotal(0) = 0
								adTotal(1) = 0
								adTotal(2) = 0
								lURCounter = 0
								adURTotal = Split("0,0,0", ",")
								adURTotal(0) = 0
								adURTotal(1) = 0
								adURTotal(2) = 0
								lURPageCount = 0
								lCTCounter = 0
								adCTTotal = Split("0,0,0", ",")
								adCTTotal(0) = 0
								adCTTotal(1) = 0
								adCTTotal(2) = 0
								lCTPageCount = 0
								bCTFirst = True
								lPageCount = 1
								bPageHeader = False
								asConceptsP = ""
								asConceptsD = ""
								bSection = False
								bSectionSummary = False
								bFirst = True
								lTempID = -1
								lEmployeeCount=0
								Do While Not oRecordset.EOF
									sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
									If bReview Then sTempCurrent = sTempCurrent & "," & CStr(oRecordset.Fields("RecordDate").Value)
									If (Len(sContents) > 0) And ((lCurrentCompanyID <> CLng(oRecordset.Fields("CompanyID").Value)) Or (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Or (StrComp(sCurrentEmployeeID, sTempCurrent, vbBinaryCompare) <> 0)) Then
										If lTempID > -1 Then
											If lTempDeduction = 0 Then
												asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
											Else
												asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
											End If
											lTempID = -1
										End If
										If lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value) Then
											bSectionSummary = True
										End If
										sConcepts = ""
										asConceptsP = Left(asConceptsP, (Len(asConceptsP) - Len(LIST_SEPARATOR)))
										asConceptsD = Left(asConceptsD, (Len(asConceptsD) - Len(LIST_SEPARATOR)))
										asConceptsP = Split(asConceptsP, LIST_SEPARATOR, -1, vbBinaryCompare)
										asConceptsD = Split(asConceptsD, LIST_SEPARATOR, -1, vbBinaryCompare)
										iBound = Int(UBound(asConceptsP) / 2)
										jBound = Int(UBound(asConceptsD) / 2)
										If iBound > jBound Then
											iLast = iBound
										Else
											iLast = jBound
										End If
										lEmployeeConceptsCount = iLast
										For iIndex = 0 To iLast
											jIndex = iIndex
											sRowContents = ""
											If iIndex <= iBound Then
												asConceptsP(iIndex) = Split(asConceptsP(iIndex), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
												sRowContents = sRowContents & CleanStringForHTML(asConceptsP(iIndex)(0) & asConceptsP(iIndex)(1)) & " "
												sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsP(iIndex)(2)), 2, True, False, True)), Len("0000000000")) & " "
													asConceptsP(iIndex)(3) = Split(asConceptsP(iIndex)(3), ".")
													sRowContents = sRowContents & asConceptsP(iIndex)(3)(0) & "-" & asConceptsP(iIndex)(3)(1)
											Else
												sRowContents = sRowContents & "                                 "
											End If
											sRowContents = sRowContents & "  "
											If (iIndex + iBound + 1) <= UBound(asConceptsP) Then
												asConceptsP(iIndex + iBound + 1) = Split(asConceptsP(iIndex + iBound + 1), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
												sRowContents = sRowContents & CleanStringForHTML(asConceptsP(iIndex + iBound + 1)(0) & asConceptsP(iIndex + iBound + 1)(1)) & " "
												sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsP(iIndex + iBound + 1)(2)), 2, True, False, True)), Len("0000000000")) & " " 
													asConceptsP(iIndex + iBound + 1)(3) = Split(asConceptsP(iIndex + iBound + 1)(3), ".")
													sRowContents = sRowContents & asConceptsP(iIndex + iBound + 1)(3)(0) & "-" & asConceptsP(iIndex + iBound + 1)(3)(1)
											Else
												sRowContents = sRowContents & "                                 "
											End If

											sRowContents = sRowContents & "           "

											If jIndex <= jBound Then
												asConceptsD(jIndex) = Split(asConceptsD(jIndex), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
												sRowContents = sRowContents & CleanStringForHTML(asConceptsD(jIndex)(0) & asConceptsD(jIndex)(1)) & " "
												sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsD(jIndex)(2)), 2, True, False, True)), Len("0000000000")) & " " 
													asConceptsD(jIndex)(3) = Split(asConceptsD(jIndex)(3), ".")
													sRowContents = sRowContents & asConceptsD(jIndex)(3)(0) & "-" & asConceptsD(jIndex)(3)(1)
											Else
												sRowContents = sRowContents & "                                 "
											End If
											sRowContents = sRowContents & "  "
											If (jIndex + jBound + 1) <= UBound(asConceptsD) Then
												asConceptsD(jIndex + jBound + 1) = Split(asConceptsD(jIndex + jBound + 1), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
												sRowContents = sRowContents & CleanStringForHTML(asConceptsD(jIndex + jBound + 1)(0) & asConceptsD(jIndex + jBound + 1)(1)) & " "
												sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsD(jIndex + jBound + 1)(2)), 2, True, False, True)), Len("0000000000")) & " " 
													asConceptsD(jIndex + jBound + 1)(3) = Split(asConceptsD(jIndex + jBound + 1)(3), ".")
													sRowContents = sRowContents & asConceptsD(jIndex + jBound + 1)(3)(0) & "-" & asConceptsD(jIndex + jBound + 1)(3)(1)
											Else
												sRowContents = sRowContents & "                                 "
											End If
											sConcepts = sConcepts & sRowContents & vbNewLine
										Next
										sContents = Replace(sContents, "<CONCEPTS />", sConcepts)
										dTemp = 0
										For iIndex = 0 To UBound(asConceptsP)
											dTemp = dTemp + CDbl(asConceptsP(iIndex)(2))
										Next
										sContents = Replace(sContents, "<PERCEPTIONS_R />", FormatNumber(dTemp, 2, True, False, True))
										sContents = Replace(sContents, "<PERCEPTIONS />", FormatNumber(dTemp, 2, True, False, True))
										dTemp = 0
										For iIndex = 0 To UBound(asConceptsD)
											dTemp = dTemp + CDbl(asConceptsD(iIndex)(2))
										Next
										sContents = Replace(sContents, "<DEDUCTIONS_R />", FormatNumber(dTemp, 2, True, False, True))
										sContents = Replace(sContents, "<DEDUCTIONS />", FormatNumber(dTemp, 2, True, False, True))
										dTemp = 0
										For iIndex = 0 To UBound(asConceptsP)
											dTemp = dTemp + CDbl(asConceptsP(iIndex)(2))
										Next
										For iIndex = 0 To UBound(asConceptsD)
											dTemp = dTemp - CDbl(asConceptsD(iIndex)(2))
										Next
										sContents = Replace(sContents, "<TOTAL_R />", FormatNumber(dTemp, 2, True, False, True))
										sContents = Replace(sContents, "<TOTAL />", FormatNumber(dTemp, 2, True, False, True))
										sContents = Replace(sContents, "<TAG />", sTags)
										sContents = Replace(sContents, "<CODES />", sCodes)
										sTags = ""
										sCodes = ""
										sEmployeeData = sContents ' Guardar el contenido del registro de empleado para agregarlo más adelante
										If bSection And (Not bFirst) Then
											lPageCount = lPageCount + 1 'Incrementa página
											lURPageCount = lURPageCount + 1
											lCTPageCount = lCTPageCount + 1
											lPageBreak = N_ROWS_PER_PAGE - lRowsCountPerPage
											For i = 0 To lPageBreak ' Rellenar con espacios en blanco las líneas para el cambio de página
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<BR />", sErrorDescription)
											Next
											bSection = False
											'bSectionSummary = True
											lRowsCountPerPage = 0
											If lEmployeeCount > N_EMPLOYEES_PER_FILE Then
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "</PRE></FONT>", sErrorDescription)
												iFileCount = iFileCount + 1
												'sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
												sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<FONT FACE=""Courier"" SIZE=""3""><PRE>", sErrorDescription)
												lEmployeeCount = 0
											End If
										End If
										' Para el total de líneas utilizadas en el registro del empleado se considera encabezado (si es por cambio de sección), datos generales y conceptos
										If bSectionSummary Then
											lRowsCountPerEmployee = lHeaderCount * (N_ROWS_PER_HEADER) + N_ROWS_FOR_EMPLOYEE_DATA + (lEmployeeConceptsCount + 1) + N_ROWS_FOR_SECTION_SUMMARY
										Else
											lRowsCountPerEmployee = lHeaderCount * (N_ROWS_PER_HEADER) + N_ROWS_FOR_EMPLOYEE_DATA + (lEmployeeConceptsCount + 1)
										End If
										If lRowsCountPerPage + lRowsCountPerEmployee < N_ROWS_PER_PAGE Then ' Validar que las líneas a agregar quepan en la pagina
											' Pintar encabezado en esta parte, si es por cambio de sección
											If lHeaderCount > 0 Then ' Si se agrego encabezado por cambio de sección, agregarlo e inicializar contador de encabezado
												sHeaderDataForEmployee = Replace(sOriginalHeaderData, "<PAGE_NUMBER />", CleanStringForHTML(lPageCount))
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sHeaderDataForEmployee, sErrorDescription)
												lHeaderCount = 0
											End If
											' Pinta Datos empleado (incluye conceptos)
											lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sEmployeeData, sErrorDescription)
											If bSectionSummary Then'And (Not bFirst) Then
												If bCTFirst Then
													lCTPageCount = lCTPageCount + 1
													bCTFirst = False
												End If
												sSectionSummaryData = vbNewLine & "REGISTROS " & FormatNumber(lCTCounter, 0, True, False, True) & "       TOTAL PERCEPCIONES " & FormatNumber(adCTTotal(1), 2, True, False, True) & "       DEDUCCIONES " & FormatNumber(adCTTotal(2), 2, True, False, True) & "       NETO " & FormatNumber(adCTTotal(0), 2, True, False, True) & "         NO.PAGINAS " & FormatNumber(lCTPageCount, 0, True, False, True) '& vbNewLine
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sSectionSummaryData, sErrorDescription)
												bSectionSummary = False
											End If
											lRowsCountPerPage = lRowsCountPerPage + lRowsCountPerEmployee
											lEmployeeCount = lEmployeeCount + 1
										Else ' Verificar el caso cuando el total sea = 52
											lPageBreak = N_ROWS_PER_PAGE - lRowsCountPerPage
											For i = 0 To lPageBreak ' Rellenar con espacios en blanco las líneas para el cambio de página
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<BR />", sErrorDescription)
												'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "*", sErrorDescription)
											Next
											If lEmployeeCount > N_EMPLOYEES_PER_FILE Then
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "</PRE></FONT>", sErrorDescription)
												iFileCount = iFileCount + 1
												'sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
												sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<FONT FACE=""Courier"" SIZE=""3""><PRE>", sErrorDescription)
												lEmployeeCount = 0
											End If
											' Pinta Encabezado
											lPageCount = lPageCount + 1 'Incrementa página
											lURPageCount = lURPageCount + 1
											lCTPageCount = lCTPageCount + 1
											If lHeaderCount = 0 Then ' Si no hubo encabezado por cambio de sección utilizar el anterior puesto que página nueva
												lRowsCountPerEmployee = lRowsCountPerEmployee + N_ROWS_PER_HEADER ' Agregar líneas del encabezado al contador de líneas del registro de empleado
											End If
											sHeaderDataForEmployee = Replace(sOriginalHeaderData, "<PAGE_NUMBER />", CleanStringForHTML(lPageCount))
											lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sHeaderDataForEmployee, sErrorDescription)
											' Pinta Datos empleado (incluye conceptos)
											lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sEmployeeData, sErrorDescription)
											If bSectionSummary And (Not bFirst) Then
												If bCTFirst Then
													lCTPageCount = lCTPageCount + 1
													bCTFirst = False
												End If
												sSectionSummaryData = vbNewLine & "REGISTROS " & FormatNumber(lCTCounter, 0, True, False, True) & "       TOTAL PERCEPCIONES " & FormatNumber(adCTTotal(1), 2, True, False, True) & "       DEDUCCIONES " & FormatNumber(adCTTotal(2), 2, True, False, True) & "       NETO " & FormatNumber(adCTTotal(0), 2, True, False, True) & "         NO.PAGINAS " & FormatNumber(lCTPageCount, 0, True, False, True) '& vbNewLine
												lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sSectionSummaryData, sErrorDescription)
												bSectionSummary = False
											End If
											lRowsCountPerPage = lRowsCountPerEmployee ' Acumular líneas utilizadas en el empleado
											'Re-inicializar variables para el registro del nuevo empleado
											sOriginalHeaderData = ""
											sEmployeeData = ""
											lHeaderCount = 0
											bPageHeader = True
										End If
										If bFirst Then
											bFirst = False
											bSection = False
										End If
									End If
									If (lCurrentCompanyID <> CLng(oRecordset.Fields("CompanyID").Value)) Or (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
										sContents = sGeneralHeader
										sContents = Replace(sContents, "<COMPANY_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value)))
										sContents = Replace(sContents, "<COMPANY_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CompanyName").Value), " ", 17, 1)))
										asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
										sContents = Replace(sContents, "<STATE_NAME />", asStateNames(CInt(asPath(2))))
										sContents = Replace(sContents, "<PAYMENT_CENTER_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)))
										sContents = Replace(sContents, "<PAYMENT_CENTER_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("PaymentCenterName").Value), " ", 80, 1)))
										sContents = Replace(sContents, "<START_PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lStartPayrollDate))
										sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
										sContents = Replace(sContents, "<ISSUE_DATE />", DisplayNumericDateFromSerialNumber(lIssueDate))
										sContents = Replace(sContents, "<ECONOMIC_ZONE_ID />", CStr(oRecordset.Fields("EconomicZoneID").Value))
										lCurrentCompanyID = CLng(oRecordset.Fields("CompanyID").Value)
										lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
										sOriginalHeaderData = sContents
										lHeaderCount = lHeaderCount + 1
										bSection = True
										bSectionSummary = False
										bPageHeader = False ' Apagar para ya no solicitar encabezado puesto que ya se tiene el del cambio de sección
										lURCounter = 0
										adURTotal = Split("0,0,0", ",")
										adURTotal(0) = 0
										adURTotal(1) = 0
										adURTotal(2) = 0
										lURPageCount = 0
										lCTCounter = 0
										adCTTotal = Split("0,0,0", ",")
										adCTTotal(0) = 0
										adCTTotal(1) = 0
										adCTTotal(2) = 0
										lCTPageCount = 0
									End If
									sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
									If bReview Then sTempCurrent = sTempCurrent & "," & CStr(oRecordset.Fields("RecordDate").Value)
									If StrComp(sCurrentEmployeeID, sTempCurrent, vbBinaryCompare) <> 0 Then
										If bPageHeader Then ' Validar si ya se tiene un encabezado por cambio de sección
											sContents = sGeneralHeader
											sContents = Replace(sContents, "<COMPANY_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value)))
											sContents = Replace(sContents, "<COMPANY_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CompanyName").Value), " ", 17, 1)))
											asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
											sContents = Replace(sContents, "<STATE_NAME />", asStateNames(CInt(asPath(2))))
											sContents = Replace(sContents, "<PAYMENT_CENTER_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)))
											sContents = Replace(sContents, "<PAYMENT_CENTER_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("PaymentCenterName").Value), " ", 80, 1, 1)))
											sContents = Replace(sContents, "<START_PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lStartPayrollDate))
											sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
											sContents = Replace(sContents, "<ISSUE_DATE />", DisplayNumericDateFromSerialNumber(lIssueDate))
											sContents = Replace(sContents, "<ECONOMIC_ZONE_ID />", CStr(oRecordset.Fields("EconomicZoneID").Value))
											lCurrentCompanyID = CLng(oRecordset.Fields("CompanyID").Value)
											lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
											sOriginalHeaderData = sContents
											bPageHeader = False
										End If
										sContents = sEmployeeHeader
										If bReview And (CLng(oRecordset.Fields("RecordDate").Value) = CLng(lForPayrollID)) Then sContents = sEmployeeHeaderTotals
										sContents = Replace(sContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)))
										Else
											sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)))
										End If
										sContents = Replace(sContents, "<JOB_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)))
										sContents = Replace(sContents, "<EMPLOYEE_START_DATE />", DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)))
										sContents = Replace(sContents, "<POSITION_SHORT_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 7, 1)))
										If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
											sContents = Replace(sContents, "<LEVEL_SHORT_NAME />", Left(Right(("000" & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)), Len("000")), Len("000")))
											sContents = Replace(sContents, "<SUBLEVEL_SHORT_NAME />", CStr(oRecordset.Fields("IntegrationID").Value))
										Else
											sContents = Replace(sContents, "<LEVEL_SHORT_NAME />", Left(Right(("00" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000")), Len("00")))
											sContents = Replace(sContents, "<SUBLEVEL_SHORT_NAME />", Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0")))
										End If
										sContents = Replace(sContents, "<RECORD_DATE />", DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RecordDate").Value)))
										sContents = Replace(sContents, "<CHECK_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("CheckNumber").Value)))
										If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
											sContents = Replace(sContents, "<PAYMENT_TYPE />", "cheque")
											sContents = Replace(sContents, "<BENEFICIARY_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)))
											sContents = Replace(sContents, "<BENEFICIARY_FULL_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("BeneficiaryLastName").Value) & " " & CStr(oRecordset.Fields("BeneficiaryLastName2").Value) & " " & CStr(oRecordset.Fields("BeneficiaryName").Value), " ", 70, 1)))
										ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
											sContents = Replace(sContents, "<PAYMENT_TYPE />", "cheque")
											sContents = Replace(sContents, "<CREDITOR_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("CreditorNumber").Value)))
											sContents = Replace(sContents, "<CREDITOR_FULL_NAME />", CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CreditorLastName").Value) & " " & CStr(oRecordset.Fields("CreditorLastName2").Value) & " " & CStr(oRecordset.Fields("CreditorName").Value), " ", 70, 1)))
										Else
											If StrComp(CStr(oRecordset.Fields("AccountNumber").Value), ".", vbBinaryCompare) = 0 Then
												sContents = Replace(sContents, "<PAYMENT_TYPE />", "cheque")
											Else
												sContents = Replace(sContents, "<PAYMENT_TYPE />", "depósito")
											End If
										End If

										If InStr(1, ",0131,0228,0229,0331,0430,0531,0630,0731,0831,0930,1031,1130,1231,", Right(CStr(oRecordset.Fields("RecordDate").Value), Len("0000")), vbBinaryCompare) > 0 Then
											If CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2)) = 1 Then
												sTags = sTags & "<37> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2)) - 1 & "  "
											Else
												sTags = sTags & "<37> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2) & "  "
											End If

											If CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2)) = 1 Then
												sTags = sTags & "<38> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2)) - 1 & "  "
											Else
												sTags = sTags & "<38> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2) & "  "
											End If

											If CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2)) = 1 Then
												sTags = sTags & "<39> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2)) - 1 & "  "
											Else
												sTags = sTags & "<39> " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2) & "  "
											End If
											sCodes = sCodes & "Cod:  " & CStr(oRecordset.Fields("Concepts40").Value) & "  "
										End If
										If InStr(1, ",0131,0430,0731,1031,", Right(CStr(oRecordset.Fields("RecordDate").Value), Len("0000")), vbBinaryCompare) > 0 Then
											If CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2)) = 1 Then
												sTags = sTags & "<40> " & asMonths(1)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2)) - 1 & " " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2)) - 1 & "  "
											Else
												sTags = sTags & "<40> " & asMonths(1)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2) & " " & asMonths(0)(CInt(Mid(CStr(oRecordset.Fields("RecordDate").Value), 5, 2))) & "/" & Mid(CStr(oRecordset.Fields("RecordDate").Value), 3, 2) & "  "
											End If
										End If

'										sPeriod = lStartPayrollDate & "-" & lForPayrollID
										sPeriod = lTempStartDate & "-" & lTempEndDate
										asConceptsP = ""
										asConceptsD = ""

										If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0 Then
											lCounter = lCounter + 1
											lURCounter = lURCounter + 1
											lCTCounter = lCTCounter + 1
										End If
										sCurrentEmployeeID = CStr(oRecordset.Fields("EmployeeID").Value)
										sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
										If bReview Then sCurrentEmployeeID = sCurrentEmployeeID & "," & CStr(oRecordset.Fields("RecordDate").Value)
									End If
									Select Case CLng(oRecordset.Fields("ConceptID").Value)
										Case -2
											sContents = Replace(sContents, "<DEDUCTIONS />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
											adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(2) = adURTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(2) = adCTTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case -1
											sContents = Replace(sContents, "<PERCEPTIONS />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
											adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(1) = adURTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(1) = adCTTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case 0
											sContents = Replace(sContents, "<TOTAL />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
											adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(0) = adURTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(0) = adCTTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case Else
											If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
												If False Then 'CLng(oRecordset.Fields("RecordDate").Value) = CLng(lPayrollID) Then
													If lTempID > -1 Then
														If lTempDeduction = 0 Then
															asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
														Else
															asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
														End If
														lTempID = -1
													End If
													asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptShortName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
													lTempID = -1
												Else
													If lTempID <> CLng(oRecordset.Fields("ConceptID").Value) Then
														If lTempID > -1 Then
															If lTempDeduction = 0 Then
																asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
															Else
																asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
															End If
														End If
														lTempID = CLng(oRecordset.Fields("ConceptID").Value)
														lTempDeduction = 0
														sTempShortName = CStr(oRecordset.Fields("ConceptShortName").Value)
														dTempAmount = 0
														lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)
														If lTempStartDate = 0 Then lTempStartDate = lPayrollID
													End If
													dTempAmount = dTempAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
													lTempEndDate = CLng(oRecordset.Fields("LastDate").Value)
													If lTempEndDate = 0 Then lTempEndDate = lPayrollID
												End If
												If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
													adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adURTotal(0) = adURTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adURTotal(1) = adURTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adCTTotal(0) = adCTTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adCTTotal(1) = adCTTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
												ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
													adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adURTotal(0) = adURTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adURTotal(1) = adURTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adCTTotal(0) = adCTTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
													adCTTotal(1) = adCTTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
												End If
											Else
												If False Then 'CLng(oRecordset.Fields("RecordDate").Value) = CLng(lPayrollID) Then
													If lTempID > -1 Then
														If lTempDeduction = 0 Then
															asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
														Else
															asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
														End If
														lTempID = -1
													End If
													asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptShortName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
													lTempID = -1
												Else
													If lTempID <> CLng(oRecordset.Fields("ConceptID").Value) Then
														If lTempID > -1 Then
															If lTempDeduction = 0 Then
																asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
															Else
																asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
															End If
														End If
														lTempID = CLng(oRecordset.Fields("ConceptID").Value)
														lTempDeduction = 1
														sTempShortName = CStr(oRecordset.Fields("ConceptShortName").Value)
														dTempAmount = 0
														lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)
														If lTempStartDate = 0 Then lTempStartDate = lPayrollID
													End If
													dTempAmount = dTempAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
													lTempEndDate = CLng(oRecordset.Fields("LastDate").Value)
													If lTempEndDate = 0 Then lTempEndDate = lPayrollID
												End If
											End If
									End Select

									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								oRecordset.Close
								If lTempID > -1 Then
									If lTempDeduction = 0 Then
										asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
									Else
										asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
									End If
									lTempID = -1
								End If
								sConcepts = ""
								asConceptsP = Left(asConceptsP, (Len(asConceptsP) - Len(LIST_SEPARATOR)))
								asConceptsD = Left(asConceptsD, (Len(asConceptsD) - Len(LIST_SEPARATOR)))
								asConceptsP = Split(asConceptsP, LIST_SEPARATOR, -1, vbBinaryCompare)
								asConceptsD = Split(asConceptsD, LIST_SEPARATOR, -1, vbBinaryCompare)
								iBound = Int(UBound(asConceptsP) / 2)
								jBound = Int(UBound(asConceptsD) / 2)
								If iBound > jBound Then
									iLast = iBound
								Else
									iLast = jBound
								End If
								For iIndex = 0 To iLast
									jIndex = iIndex
									sRowContents = ""
									If iIndex <= iBound Then
										asConceptsP(iIndex) = Split(asConceptsP(iIndex), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
										sRowContents = sRowContents & CleanStringForHTML(asConceptsP(iIndex)(0) & asConceptsP(iIndex)(1)) & " "
										sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsP(iIndex)(2)), 2, True, False, True)), Len("0000000000")) & " "
											asConceptsP(iIndex)(3) = Split(asConceptsP(iIndex)(3), ".")
											sRowContents = sRowContents & asConceptsP(iIndex)(3)(0) & "-" & asConceptsP(iIndex)(3)(1)
									Else
										sRowContents = sRowContents & "                                 "
									End If
									sRowContents = sRowContents & "  "
									If (iIndex + iBound + 1) <= UBound(asConceptsP) Then
										asConceptsP(iIndex + iBound + 1) = Split(asConceptsP(iIndex + iBound + 1), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
										sRowContents = sRowContents & CleanStringForHTML(asConceptsP(iIndex + iBound + 1)(0) & asConceptsP(iIndex + iBound + 1)(1)) & " "
										sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsP(iIndex + iBound + 1)(2)), 2, True, False, True)), Len("0000000000")) & " " 
											asConceptsP(iIndex + iBound + 1)(3) = Split(asConceptsP(iIndex + iBound + 1)(3), ".")
											sRowContents = sRowContents & asConceptsP(iIndex + iBound + 1)(3)(0) & "-" & asConceptsP(iIndex + iBound + 1)(3)(1)
									Else
										sRowContents = sRowContents & "                                 "
									End If
									sRowContents = sRowContents & "           "
									If jIndex <= jBound Then
										asConceptsD(jIndex) = Split(asConceptsD(jIndex), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
										sRowContents = sRowContents & CleanStringForHTML(asConceptsD(jIndex)(0) & asConceptsD(jIndex)(1)) & " "
										sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsD(jIndex)(2)), 2, True, False, True)), Len("0000000000")) & " " 
											asConceptsD(jIndex)(3) = Split(asConceptsD(jIndex)(3), ".")
											sRowContents = sRowContents & asConceptsD(jIndex)(3)(0) & "-" & asConceptsD(jIndex)(3)(1)
									Else
										sRowContents = sRowContents & "                                 "
									End If
									sRowContents = sRowContents & "  "
									If (jIndex + jBound + 1) <= UBound(asConceptsD) Then
										asConceptsD(jIndex + jBound + 1) = Split(asConceptsD(jIndex + jBound + 1), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
										sRowContents = sRowContents & CleanStringForHTML(asConceptsD(jIndex + jBound + 1)(0) & asConceptsD(jIndex + jBound + 1)(1)) & " "
										sRowContents = sRowContents & Right(("          " & FormatNumber(CDbl(asConceptsD(jIndex + jBound + 1)(2)), 2, True, False, True)), Len("0000000000")) & " " 
											asConceptsD(jIndex + jBound + 1)(3) = Split(asConceptsD(jIndex + jBound + 1)(3), ".")
											sRowContents = sRowContents & asConceptsD(jIndex + jBound + 1)(3)(0) & "-" & asConceptsD(jIndex + jBound + 1)(3)(1)
									Else
										sRowContents = sRowContents & "                                 "
									End If
									sConcepts = sConcepts & sRowContents & vbNewLine
								Next
								sContents = Replace(sContents, "<CONCEPTS />", sConcepts)
								dTemp = 0
								For iIndex = 0 To UBound(asConceptsP)
									dTemp = dTemp + CDbl(asConceptsP(iIndex)(2))
								Next
								sContents = Replace(sContents, "<PERCEPTIONS_R />", FormatNumber(dTemp, 2, True, False, True))
								sContents = Replace(sContents, "<PERCEPTIONS />", FormatNumber(dTemp, 2, True, False, True))
								dTemp = 0
								For iIndex = 0 To UBound(asConceptsD)
									dTemp = dTemp + CDbl(asConceptsD(iIndex)(2))
								Next
								sContents = Replace(sContents, "<DEDUCTIONS_R />", FormatNumber(dTemp, 2, True, False, True))
								sContents = Replace(sContents, "<DEDUCTIONS />", FormatNumber(dTemp, 2, True, False, True))
								dTemp = 0
								For iIndex = 0 To UBound(asConceptsP)
									dTemp = dTemp + CDbl(asConceptsP(iIndex)(2))
								Next
								For iIndex = 0 To UBound(asConceptsD)
									dTemp = dTemp - CDbl(asConceptsD(iIndex)(2))
								Next
								sContents = Replace(sContents, "<TOTAL_R />", FormatNumber(dTemp, 2, True, False, True))
								sContents = Replace(sContents, "<TOTAL />", FormatNumber(dTemp, 2, True, False, True))
								sContents = Replace(sContents, "<TAG />", sTags)
								sContents = Replace(sContents, "<CODES />", sCodes)
								sTags = ""
								sCodes = ""
								sEmployeeData = sContents ' Guardar el contenido del registro de empleado para agregarlo más adelante
								If bSection and (Not bFirst) Then
									lPageCount = lPageCount + 1 'Incrementa página
									lURPageCount = lURPageCount + 1
									lCTPageCount = lCTPageCount + 1
									lPageBreak = N_ROWS_PER_PAGE - lRowsCountPerPage
									For i = 0 To lPageBreak ' Rellenar con espacios en blanco las líneas para el cambio de página
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<BR />", sErrorDescription)
									Next
									bSection = False
									lRowsCountPerPage = 0
									If lEmployeeCount > N_EMPLOYEES_PER_FILE Then
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "</PRE></FONT>", sErrorDescription)
										iFileCount = iFileCount + 1
										'sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
										sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<FONT FACE=""Courier"" SIZE=""3""><PRE>", sErrorDescription)
										lEmployeeCount = 0
									End If
								End If
								lRowsCountPerEmployee = lHeaderCount * (N_ROWS_PER_HEADER) + N_ROWS_FOR_EMPLOYEE_DATA + (lEmployeeConceptsCount + 1) + N_ROWS_FOR_SECTION_SUMMARY
								If lRowsCountPerPage + lRowsCountPerEmployee < N_ROWS_PER_PAGE Then ' Validar que las líneas a agregar quepan en la pagina
									' Pintar encabezado en esta parte, si es por cambio de sección
									If lHeaderCount > 0 Then ' Si se agrego encabezado por cambio de sección, agregarlo e inicializar contador de encabezado
										sHeaderDataForEmployee = Replace(sOriginalHeaderData, "<PAGE_NUMBER />", CleanStringForHTML(lPageCount))
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sHeaderDataForEmployee, sErrorDescription)
										lHeaderCount = 0
									End If
									' Pinta Datos empleado (incluye conceptos)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sEmployeeData, sErrorDescription)
									If bCTFirst Then
										lCTPageCount = lCTPageCount + 1
										bCTFirst = False
									End If
									sSectionSummaryData = vbNewLine & "REGISTROS " & FormatNumber(lCTCounter, 0, True, False, True) & "       TOTAL PERCEPCIONES " & FormatNumber(adCTTotal(1), 2, True, False, True) & "       DEDUCCIONES " & FormatNumber(adCTTotal(2), 2, True, False, True) & "       NETO " & FormatNumber(adCTTotal(0), 2, True, False, True) & "         NO.PAGINAS " & FormatNumber(lCTPageCount, 0, True, False, True) '& vbNewLine
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sSectionSummaryData, sErrorDescription)
									lRowsCountPerPage = lRowsCountPerPage + lRowsCountPerEmployee
									lEmployeeCount = lEmployeeCount + 1
								Else ' Verificar el caso cuando el total sea = 52
									lPageBreak = N_ROWS_PER_PAGE - lRowsCountPerPage
									For i = 0 To lPageBreak ' Rellenar con espacios en blanco las líneas para el cambio de página
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<BR />", sErrorDescription)
									Next
									If lEmployeeCount > N_EMPLOYEES_PER_FILE Then
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "</PRE></FONT>", sErrorDescription)
										iFileCount = iFileCount + 1
										'sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
										sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & iFileCount
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "<FONT FACE=""Courier"" SIZE=""3""><PRE>", sErrorDescription)
										lEmployeeCount = 0
									End If
									' Pinta Encabezado
									lPageCount = lPageCount + 1 'Incrementa página
									lURPageCount = lURPageCount + 1
									lCTPageCount = lCTPageCount + 1
									If lHeaderCount = 0 Then ' Si no hubo encabezado por cambio de sección utilizar el anterior puesto que página nueva
										lRowsCountPerEmployee = lRowsCountPerEmployee + N_ROWS_PER_HEADER ' Agregar líneas del encabezado al contador de líneas del registro de empleado
									End If
									sHeaderDataForEmployee = Replace(sOriginalHeaderData, "<PAGE_NUMBER />", CleanStringForHTML(lPageCount))
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sHeaderDataForEmployee, sErrorDescription)
									' Pinta Datos empleado (incluye conceptos)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sEmployeeData, sErrorDescription)
									If bCTFirst Then
										lCTPageCount = lCTPageCount + 1
										bCTFirst = False
									End If
									sSectionSummaryData = vbNewLine & "REGISTROS " & FormatNumber(lCTCounter, 0, True, False, True) & "       TOTAL PERCEPCIONES " & FormatNumber(adCTTotal(1), 2, True, False, True) & "       DEDUCCIONES " & FormatNumber(adCTTotal(2), 2, True, False, True) & "       NETO " & FormatNumber(adCTTotal(0), 2, True, False, True) & "         NO.PAGINAS " & FormatNumber(lCTPageCount, 0, True, False, True) '& vbNewLine
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sSectionSummaryData, sErrorDescription)
									lRowsCountPerPage = lRowsCountPerEmployee ' Acumular líneas utilizadas en el empleado
									'Re-inicializar variables para el registro del nuevo empleado
									sOriginalHeaderData = ""
									sEmployeeData = ""
									lHeaderCount = 0
									bPageHeader = True
								End If
								sContents = vbNewLine & vbNewLine & vbNewLine
								sContents = sContents & "TOTAL DE REGISTROS " & FormatNumber(lCounter, 0, True, False, True) & "       TOTAL PERCEPCIONES " & FormatNumber(adTotal(1), 2, True, False, True) & "       DEDUCCIONES " & FormatNumber(adTotal(2), 2, True, False, True) & "       NETO " & FormatNumber(adTotal(0), 2, True, False, True) & "         NO.PAGINAS " & FormatNumber(lPageCount, 0, True, False, True)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), sContents, sErrorDescription)
							lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".htm"), "</PRE></FONT>", sErrorDescription)
							lErrorNumber = ZipFile(Server.MapPath(sFolderPath), Server.MapPath(sZipFile), sErrorDescription)
							If lErrorNumber = 0 Then
								Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
								sErrorDescription = "No se pudieron guardar la información del reporte."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = DeleteFolder(Server.MapPath(sFolderPath), sErrorDescription)
							End If
							oEndDate = Now()
							If (lErrorNumber = 0) And B_USE_SMTP Then
								If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
							End If
						End If
					End If
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReports1003 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1004(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Resumen por conceptos de nómina. Reporte basado en la
'         hoja 001128 Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bShowTotals, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1004"
	Dim sHeaderContents
	Dim sCondition
	Dim sDistinct
	Dim iMonth
	Dim iYear
	Dim iLastDay
	Dim lPayments
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "PaymentCenters.AreaID", "EmployeesHistoryList.PaymentCenterID")
	iMonth = CInt(oRequest("MonthID").Item)
	iYear = CInt(oRequest("YearID").Item)
	Select Case iMonth
		Case 1,3,5,7,8,10,12
			iLastDay = 31
		Case 2
			If (iYear Mod 4) = 0 Then
				iLastDay = 29
			Else
				iLastDay = 28
			End If
		Case 4,6,9,11
			iLastDay = 30
	End Select
	If (iConnectionType <> ACCESS) And (iConnectionType <> ACCESS_DSN) Then
		sDistinct = "Distinct "
	Else
		sDistinct = ""
	End If

	sErrorDescription = "No se pudieron obtener los conceptos de pago."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select IsDeduction, ConceptShortName, ConceptName, Partida.BudgetShortName As PartidaShortName, Subpartida.BudgetShortName As SubpartidaShortName, Budgets.BudgetShortName, Count(" & sDistinct & "Payroll_" & iYear & ".EmployeeID) As TotalCount, Sum(Payroll_" & iYear & ".ConceptAmount) As TotalAmount From Payroll_" & iYear & ", EmployeesChangesLKP, EmployeesHistoryList, BankAccounts, Areas, Zones, Concepts, Budgets As Partida, Budgets As Subpartida, Budgets Where (Payroll_" & iYear & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & iYear & iMonth & iLastDay & ") And (Payroll_" & iYear & ".RecordDate>=" & iYear & iMonth & "00) And (Payroll_" & iYear & ".RecordDate<=" & iYear & iMonth & "99) And (BankAccounts.StartDate<=" & iYear & iMonth & iLastDay & ") And (BankAccounts.EndDate>=" & iYear & iMonth & iLastDay & ") And (BankAccounts.Active=1) And (Areas.StartDate<=" & iYear & iMonth & iLastDay & ") And (Areas.EndDate>=" & iYear & iMonth & iLastDay & ") And (Concepts.StartDate<=" & iYear & iMonth & iLastDay & ") And (Concepts.EndDate>=" & iYear & iMonth & iLastDay & ") And (Payroll_" & iYear & ".ConceptID=Concepts.ConceptID) And (Partida.BudgetID=Subpartida.ParentID) And (Subpartida.BudgetID=Budgets.ParentID) And (Budgets.BudgetID=Concepts.BudgetID) And (Concepts.ConceptID>0) " & sCondition & " Group By IsDeduction, ConceptShortName, ConceptName, Partida.BudgetShortName, Subpartida.BudgetShortName, Budgets.BudgetShortName Order By IsDeduction, ConceptShortName", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1007.htm"), sErrorDescription)
			sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
			Response.Write sHeaderContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Concepto,Partida,Subpartida,Tipo,Empleados,Descripción", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,300", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",CENTER,CENTER,CENTER,RIGHT,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PartidaShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubpartidaShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("TotalCount").Value), 0, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop

				sErrorDescription = "No se pudieron obtener los conceptos de pago."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select IsDeduction, ConceptName, Count(" & sDistinct & "Payroll_" & iYear & ".EmployeeID) As TotalCount, Sum(Payroll_" & iYear & ".ConceptAmount) As TotalAmount From Payroll_" & iYear & ", EmployeesChangesLKP, EmployeesHistoryList, BankAccounts, Areas, Zones, Concepts Where (Payroll_" & iYear & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & iYear & iMonth & iLastDay & ") And (Payroll_" & iYear & ".RecordDate>=" & iYear & iMonth & "00) And (Payroll_" & iYear & ".RecordDate<=" & iYear & iMonth & "99) And (BankAccounts.StartDate<=" & iYear & iMonth & iLastDay & ") And (BankAccounts.EndDate>=" & iYear & iMonth & iLastDay & ") And (BankAccounts.Active=1) And (Areas.StartDate<=" & iYear & iMonth & iLastDay & ") And (Areas.EndDate>=" & iYear & iMonth & iLastDay & ") And (Payroll_" & iYear & ".ConceptID=Concepts.ConceptID) And (Concepts.ConceptID<=0) " & sCondition & " Group By IsDeduction, ConceptName Order By IsDeduction", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sRowContents = "<SPAN COLS=""4"" /><B>TOTAL " & UCase(CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))) & ":</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"

						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lPayments = CLng(oRecordset.Fields("TotalCount").Value)
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					sRowContents = "<SPAN COLS=""4"" /><B>TOTAL DE REGISTROS:</B>"
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lPayments, 0, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen conceptos de pago registrados en el sistema."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1004 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1004Sp(oRequest, oADODBConnection, bShowTotals, bForExport, sErrorDescription)
'************************************************************
'Purpose: Resumen por conceptos de nómina. Reporte basado en la
'         hoja 001128 Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bShowTotals, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1004Sp"
	Dim sHeaderContents
	Dim sCondition
	Dim iMonth
	Dim iYear
	Dim iIndex
	Dim asPayrollIDs
	Dim lPayments
	Dim dPayments
	Dim sQuery
	Dim oRecordset
	Dim oPayrollRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	iMonth = CInt(oRequest("MonthID").Item)
	iYear = CInt(oRequest("YearID").Item)
	sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollDate>=" & (iYear & Right(("0" & iMonth), Len("00"))) & "00) And (PayrollDate<=" & (iYear & Right(("0" & iMonth), Len("00"))) & "99)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrollIDs = ""
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asPayrollIDs = asPayrollIDs & CStr(oRecordset.Fields("PayrollID").Value) & ",0;"
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1004.htm"), sErrorDescription)
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
		Response.Write sHeaderContents

		sErrorDescription = "No se pudieron obtener los conceptos de pago."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptShortName, ConceptName, Partida.BudgetShortName As PartidaShortName, Subpartida.BudgetShortName As SubpartidaShortName, Budgets.BudgetShortName From Concepts, Budgets As Partida, Budgets As Subpartida, Budgets Where (Partida.BudgetID=Subpartida.ParentID) And (Subpartida.BudgetID=Budgets.ParentID) And (Budgets.BudgetID=Concepts.BudgetID) And (ConceptID>0) Order By ConceptShortName", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If bShowTotals Then
				Do While Not oRecordset.EOF
					lPayments = 0
					For iIndex = 0 To UBound(asPayrollIDs)
						sErrorDescription = "No se pudieron obtener los montos pagados por conceptos de pago para el periodo especificado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalPayments From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & ") " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
						If lErrorNumber = 0 Then
							If Not oPayrollRecordset.EOF Then
								If Not IsNull(oPayrollRecordset.Fields("TotalPayments").Value) Then lPayments = lPayments + CDbl(oPayrollRecordset.Fields("TotalPayments").Value)
								oPayrollRecordset.Close
							End If
						End If
					Next

					If lPayments > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PartidaShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("SubpartidaShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(lPayments, 2, True, False, True) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD>"
						Response.Write "</TR>"
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Else
				Do While Not oRecordset.EOF
					For iIndex = 0 To UBound(asPayrollIDs)
						asPayrollIDs(iIndex)(1) = 0
						sErrorDescription = "No se pudieron obtener los montos pagados por conceptos de pago para el periodo especificado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID) As TotalPayments From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & ") " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
						If lErrorNumber = 0 Then
							If Not oPayrollRecordset.EOF Then
								If Not IsNull(oPayrollRecordset.Fields("TotalPayments").Value) Then asPayrollIDs(iIndex)(1) = CLng(oPayrollRecordset.Fields("TotalPayments").Value)
								oPayrollRecordset.Close
							End If
						End If
					Next

					If UBound(asPayrollIDs) > 0 Then
						sQuery = "Select Count(Payroll_" & asPayrollIDs(0)(0) & ".EmployeeID) As TotalPayments From "
						For iIndex = 0 To UBound(asPayrollIDs)
							sQuery = sQuery & "Payroll_" & asPayrollIDs(iIndex)(0) & ", "
						Next
						sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
						sQuery = sQuery & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(0)(0) & ".ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & ")"
						For iIndex = 0 To UBound(asPayrollIDs) - 1
							sQuery = sQuery & " And (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=Payroll_" & asPayrollIDs(iIndex + 1)(0) & ".EmployeeID) And (Payroll_" & asPayrollIDs(iIndex)(0) & ".ConceptID=Payroll_" & asPayrollIDs(iIndex + 1)(0) & ".ConceptID)"
						Next
						sQuery = sQuery & " And (Payroll_" & asPayrollIDs(0)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(0)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(0)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(0)(0) & ") " & sCondition

						sErrorDescription = "No se pudo obtener el número de empleados pagados por conceptos de pago para el periodo especificado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
						If lErrorNumber = 0 Then
							If Not oPayrollRecordset.EOF Then
								lPayments = CLng(oPayrollRecordset.Fields("TotalPayments").Value)
								For iIndex = 0 To UBound(asPayrollIDs)
									If Not IsNull(oPayrollRecordset.Fields("TotalPayments").Value) Then lPayments = lPayments + (CLng(asPayrollIDs(iIndex)(1)) - CLng(oPayrollRecordset.Fields("TotalPayments").Value))
								Next
								oPayrollRecordset.Close
							End If
						End If
					Else
						lPayments = CLng(asPayrollIDs(0)(1))
					End If

					If lPayments > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PartidaShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("SubpartidaShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value)) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(lPayments, 0, True, False, True) & "</FONT></TD>"
							Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Courier"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD>"
						Response.Write "</TR>"
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If

			dPayments = 0
			For iIndex = 0 To UBound(asPayrollIDs)
				sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=-1)" & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
				If lErrorNumber = 0 Then
					If Not oPayrollRecordset.EOF Then
						dPayments = dPayments + CDbl(oPayrollRecordset.Fields("TotalAmount").Value)
						oPayrollRecordset.Close
					End If
				End If
			Next
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;</FONT></TD>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">TOTAL PERCEPCIONES:</FONT></TD>"
				Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(dPayments, 2, True, False, True) & "</FONT></TD>"
			Response.Write "</TR>"

			dPayments = 0
			For iIndex = 0 To UBound(asPayrollIDs)
				sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=-2) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
				If lErrorNumber = 0 Then
					If Not oPayrollRecordset.EOF Then
						dPayments = dPayments + CDbl(oPayrollRecordset.Fields("TotalAmount").Value)
						oPayrollRecordset.Close
					End If
				End If
			Next
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;</FONT></TD>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">TOTAL DEDUCCIONES:</FONT></TD>"
				Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(dPayments, 2, True, False, True) & "</FONT></TD>"
			Response.Write "</TR>"

			dPayments = 0
			For iIndex = 0 To UBound(asPayrollIDs)
				sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=0) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
				If lErrorNumber = 0 Then
					If Not oPayrollRecordset.EOF Then
						dPayments = dPayments + CDbl(oPayrollRecordset.Fields("TotalAmount").Value)
						oPayrollRecordset.Close
					End If
				End If
			Next
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;</FONT></TD>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">TOTAL NETO:</FONT></TD>"
				Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(dPayments, 2, True, False, True) & "</FONT></TD>"
			Response.Write "</TR>"


			For iIndex = 0 To UBound(asPayrollIDs)
				asPayrollIDs(iIndex)(1) = 0
				sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID) As TotalPayments From Payroll_" & asPayrollIDs(iIndex)(0) & ", EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & asPayrollIDs(iIndex)(0) & ") And (Areas.StartDate<=" & asPayrollIDs(iIndex)(0) & ") And (Areas.EndDate>=" & asPayrollIDs(iIndex)(0) & ") And (ConceptID=0) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
				If lErrorNumber = 0 Then
					If Not oPayrollRecordset.EOF Then
						If Not IsNull(oPayrollRecordset.Fields("TotalPayments").Value) Then asPayrollIDs(iIndex)(1) = CLng(oPayrollRecordset.Fields("TotalPayments").Value)
						oPayrollRecordset.Close
					End If
				End If
			Next
			If UBound(asPayrollIDs) > 0 Then
				sQuery = "Select Count(Payroll_" & asPayrollIDs(0)(0) & ".EmployeeID) As TotalPayments From "
				For iIndex = 0 To UBound(asPayrollIDs)
					sQuery = sQuery & "Payroll_" & asPayrollIDs(iIndex)(0) & ", "
				Next
				sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
				sQuery = sQuery & " Where (Payroll_" & asPayrollIDs(0)(0) & ".ConceptID=0)"
				For iIndex = 0 To UBound(asPayrollIDs) - 1
					sQuery = sQuery & " And (Payroll_" & asPayrollIDs(iIndex)(0) & ".EmployeeID=Payroll_" & asPayrollIDs(iIndex + 1)(0) & ".EmployeeID) And (Payroll_" & asPayrollIDs(iIndex)(0) & ".ConceptID=Payroll_" & asPayrollIDs(iIndex + 1)(0) & ".ConceptID)"
				Next

				sErrorDescription = "No se pudo obtener el número de empleados pagados por conceptos de pago para el periodo especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
				If lErrorNumber = 0 Then
					If Not oPayrollRecordset.EOF Then
						lPayments = CLng(oPayrollRecordset.Fields("TotalPayments").Value)
						For iIndex = 0 To UBound(asPayrollIDs)
							If Not IsNull(oPayrollRecordset.Fields("TotalPayments").Value) Then lPayments = lPayments + (CLng(asPayrollIDs(iIndex)(1)) - CLng(oPayrollRecordset.Fields("TotalPayments").Value))
						Next
						oPayrollRecordset.Close
					End If
				End If
			Else
				lPayments = CLng(asPayrollIDs(0)(1))
			End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">&nbsp;</FONT></TD>"
				Response.Write "<TD COLSPAN=""4""><FONT FACE=""Courier"" SIZE=""2"">TOTAL DE REGISTROS:</FONT></TD>"
				Response.Write "<TD ALIGN=""RIGHT""><FONT FACE=""Courier"" SIZE=""2"">" & FormatNumber(lPayments, 0, True, False, True) & "</FONT></TD>"
			Response.Write "</TR>"

			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen conceptos de pago registrados en el sistema."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1004Sp = lErrorNumber
	Err.Clear
End Function

Function BuildReport1005(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Hoja informativa de nómina de fecha de pago. Reporte basado en la hoja 00966
'		  Genera un archivo zip con todas las hojas informativas de nómina.
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1005"
	Dim sCondition
	Dim lReportID
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sHeaderContents
	Dim sContents
	Dim sRowContents
	Dim sTemp
	Dim dTemp
	Dim sDate
	Dim oStartDate
	Dim sFileName
	Dim lCurrentID
	Dim oRecordset
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "Jobs.", "EmployeeHistoryList.")
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	oStartDate = Now()
	sHeaderContents = ""
	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1005_RTF.htm"), sErrorDescription)
	If Len(sHeaderContents) > 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, EmployeeLastName, Case EmployeeLastName2 When NULL Then ' ' Else EmployeeLastName2 End EmployeeLastName2, EmployeeName, Case RFC When Null Then ' ' Else RFC End RFC, Employees.StartDate, Employees.EmployeeNumber, CompanyShortName, CompanyName, ParentZones.ZoneName As StateName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, EmployeeTypeShortName, PositionTypeShortName, EmployeesHistoryListForPayroll.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, RecordDate, ConceptAmount, ConceptTaxes, PaymentsMessages.Comments From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, PaymentsMessages, Companies, Zones, Zones As Zones2, Zones As ParentZones, Areas, Areas As ParentAreas, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.EmployeeID=PaymentsMessages.EmployeeID) And (PaymentsMessages.ConceptID=Concepts.ConceptID) And (PaymentsMessages.PayrollID=" & lPayrollID & ") And (PaymentsMessages.bSpecial In (1,3)) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.PaymentCenterID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Employees.EmployeeID, EmployeeLastName, Case EmployeeLastName2 When NULL Then ' ' Else EmployeeLastName2 End EmployeeLastName2, EmployeeName, Case RFC When Null Then ' ' Else RFC End , Employees.StartDate, Employees.EmployeeNumber, CompanyShortName, CompanyName, ParentZones.ZoneName As StateName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, EmployeeTypeShortName, PositionTypeShortName, EmployeesHistoryListForPayroll.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, RecordDate, ConceptAmount, ConceptTaxes, PaymentsMessages.Comments From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, PaymentsMessages, Companies, Zones, Zones As Zones2, Zones As ParentZones, Areas, Areas As ParentAreas, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.EmployeeID=PaymentsMessages.EmployeeID) And (PaymentsMessages.ConceptID=Concepts.ConceptID) And (PaymentsMessages.PayrollID=" & lPayrollID & ") And (PaymentsMessages.bSpecial In (1,3)) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.PaymentCenterID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, ConceptShortName -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sDate = GetSerialNumberForDate("")
				sFileName = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc")
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Replace(sFileName, ".doc", ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

				sContents = RTF_BEGIN_V & RTF_DEFAULT_TITLE & RTF_HEADER_BEGIN & RTF_HEADER_END & RTF_FOOTER_WITH_PAGE
				lErrorNumber = SaveTextToFile(sFileName, sContents, sErrorDescription)

				lCurrentID = -2
				dTemp = Left(lForPayrollID, Len("0000"))
				Select Case Mid(lForPayrollID, Len("00000"), Len("00"))
					Case "01"
						dTemp = (CInt(Left(lForPayrollID, Len("0000"))) - 1) & "12"
					Case Else
						dTemp = dTemp & "0" & (CInt(Mid(lForPayrollID, Len("00000"), Len("00"))) - 1)
				End Select
				dTemp = Right(dTemp, Len("YYMM"))
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID <> -2 Then
							sRowContents = sRowContents & RFT_NEW_PAGE
							lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)
						End If
						sContents = sHeaderContents
						sContents = Replace(sContents, "<COMPANY_SHORT_NAME />", CStr(oRecordset.Fields("CompanyShortName").Value))
						sContents = Replace(sContents, "<COMPANY_NAME />", SizeText(CStr(oRecordset.Fields("CompanyName").Value), " ", 17, 1))
						sContents = Replace(sContents, "<STATE_NAME />", SizeText(CStr(oRecordset.Fields("StateName").Value), " ", 19, 1))
						sContents = Replace(sContents, "<PAYMENT_CENTER_SHORT_NAME />", CStr(oRecordset.Fields("PaymentCenterShortName").Value))
						sContents = Replace(sContents, "<PAYMENT_CENTER_NAME />", SizeText(CStr(oRecordset.Fields("PaymentCenterName").Value), " ", 80, 1))
						sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
						sContents = Replace(sContents, "<EMPLOYEE_NUMBER />", SizeText(CStr(oRecordset.Fields("EmployeeNumber").Value), " ", 8, 1))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", SizeText(Trim(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & Trim(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & Trim(CStr(oRecordset.Fields("EmployeeName").Value)), " ", 27, 1))
						Else
							sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", SizeText(Trim(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & Trim(CStr(oRecordset.Fields("EmployeeName").Value)), " ", 27, 1))
						End If
						sContents = Replace(sContents, "<JOB_NUMBER />", SizeText(CStr(oRecordset.Fields("JobNumber").Value), " ", 6, 1))
						sContents = Replace(sContents, "<POSITION_SHORT_NAME />", SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 5, 1))
						sContents = Replace(sContents, "<LEVEL_SHORT_NAME />", Left(Right(("00" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000")), Len("00")))
						sContents = Replace(sContents, "<SUBLEVEL_SHORT_NAME />", Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0")))
						sContents = Replace(sContents, "<RFC />", SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 13, 1))
						sContents = Replace(sContents, "<EMPLOYEE_TYPE />", Left(Right(("00" & CStr(oRecordset.Fields("EmployeeTypeShortName").Value)), Len("000")), Len("00")))
						sContents = Replace(sContents, "<POSITION_TYPE />", Right(CStr(oRecordset.Fields("PositionTypeShortName").Value), Len("0")))
						sContents = Replace(sContents, "<EMPLOYEE_START_DATE />", DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)))
						If StrComp(oRecordset.Fields("AccountNumber").Value, ".", vbBinaryCompare) = 0 Then
							sContents = Replace(sContents, "<PAYMENT_TYPE />", "CHEQUES")
						Else
							sContents = Replace(sContents, "<PAYMENT_TYPE />", "DÉBITO")
						End If
						lErrorNumber = AppendTextToFile(sFileName, sContents, sErrorDescription)
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						sRowContents = ""
					End If
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						sRowContents = sRowContents & " " & RFT_NEW_LINE & " " & RTF_PARAGRAPH_BEGIN & " "
						sRowContents = sRowContents & CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value) & ": "
						sRowContents = sRowContents & ""
						sRowContents = sRowContents & RTF_TAB & Right(("          " & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)), Len("          "))
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 40, 41, 42
								If CInt(Mid(lForPayrollID, Len("YYYYM"), Len("MM"))) <= 2 Then
									sRowContents = sRowContents & RTF_TAB & " Estímulos correspondientes al mes de Diciembre de " & CInt(Left(lForPayrollID, Len("YYYY"))) - 1
								Else
									sRowContents = sRowContents & RTF_TAB & " Estímulos correspondientes al mes de " & asMonthNames_es(CInt(Mid(lForPayrollID, Len("YYYYM"), Len("MM"))) - 1) & " de " & Left(lForPayrollID, Len("YYYY"))
								End If
							Case 43
								If CInt(oRecordset.Fields("ConceptTaxes").Value) > 0 Then
									sRowContents = sRowContents & RTF_TAB & " Estímulos correspondientes a "

									sTemp = ""
									If StrComp(Right(("000" & CStr(oRecordset.Fields("ConceptTaxes").Value)), Len("1")), "1", vbBinaryCompare) = 0 Then
										sTemp = sTemp & asMonthNames_es(CInt(Right(dTemp, Len("MM"))) - 2) & ", "
									End If
									If StrComp(Left(Right(("000" & CStr(oRecordset.Fields("ConceptTaxes").Value)), Len("11")), Len("1")), "1", vbBinaryCompare) = 0 Then
										sTemp = sTemp & asMonthNames_es(CInt(Right(dTemp, Len("MM"))) - 1) & ", "
									End If
									If CLng(oRecordset.Fields("ConceptTaxes").Value) >= 100 Then
										sTemp = sTemp & asMonthNames_es(CInt(Right(dTemp, Len("MM")))) & ", "
									End If
									sTemp = Left(sTemp, (Len(sTemp) - Len(", ")))

									sRowContents = sRowContents & sTemp & " de "
									If CInt(Mid(lForPayrollID, Len("YYYYM"), Len("MM"))) <= 2 Then
										sRowContents = sRowContents & CInt(Left(lForPayrollID, Len("YYYY"))) - 1
									Else
										sRowContents = sRowContents & CInt(Left(lForPayrollID, Len("YYYY")))
									End If
								End If
							Case 50
								If CInt(Mid(lForPayrollID, Len("YYYYM"), Len("MM"))) <= 2 Then
									sRowContents = sRowContents & RTF_TAB & " Estímulos correspondientes al mes de Diciembre de " & CInt(Left(lForPayrollID, Len("YYYY"))) - 1
								Else
									sRowContents = sRowContents & RTF_TAB & " Estímulos correspondientes al mes de " & asMonthNames_es(CInt(Mid(lForPayrollID, Len("YYYYM"), Len("MM"))) - 1) & " de " & Left(lForPayrollID, Len("YYYY"))
								End If
							Case Else
								sRowContents = sRowContents & RTF_TAB & " " & Replace(Replace(CStr(oRecordset.Fields("Comments").Value), "Concepto " & CStr(oRecordset.Fields("ConceptShortName").Value) & ". ", ""), "TIEMPO EXTRA: ", "")
						End Select
						sRowContents = sRowContents & " " & RTF_PARAGRAPH_END
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = sRowContents & RTF_END
				lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFileName, Replace(sFileName, ".doc", ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFileName, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFileName, ".doc", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			End If
		End If
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen la plantilla del reporte."
	End If

	Set oRecordset = Nothing
	BuildReport1005 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1006(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte concentrado de conceptos de la nómina ordinaria. Reporte basado en la hoja 00970
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1006"
	Dim sDistinct
	Dim sCondition
	Dim oRecordset
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber
	Dim asRowContents
	Dim sRowContents
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim mConcepts(136,10)
	Dim sHeaderContents
	Dim iRowP
	Dim iRowD
	Dim iCount
	Dim iIndex
	Dim asCellWidths
	Dim asCellAlignments
	Dim bBorder
	Dim dTotalPerceptions
	Dim dTotalDeductions

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(sCondition, "(Areas.", "(Areas1."), "Employees.", "EmployeesHistoryList."), "Banks.", "BankAccounts.")
		If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'			sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
			sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		End If
	sDistinct = ""
	If (iConnectionType <> ACCESS) And (iConnectionType <> ACCESS_DSN) Then sDistinct = "Distinct "

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener las percepciones de los empleados registrados en el sistema."
	If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		End If
	ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		End If
	Else
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
		End If
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1006.htm"), sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				If bForExport Then
					sRowContents = RTF_BEGIN_H
					sRowContents = sRowContents & RTF_DEFAULT_TITLE
					sRowContents = sRowContents & RTF_HEADER_BEGIN
					sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN
					sRowContents = sRowContents & RTF_CENTER & RTF_BOLD
					sRowContents = sRowContents & RTF_FONT18_START
					sRowContents = sRowContents & sHeaderContents & RFT_NEW_LINE & " "
					sRowContents = sRowContents & "REPORTE CONCENTRADO DE CONCEPTOS DE LA NÓMINA ORDINARIA DE EMPLEADOS DEL: " & DisplayNumericDateFromSerialNumber(lPayrollID) & RFT_NEW_LINE & " " & RFT_NEW_LINE & " "
					sRowContents = sRowContents & asTitles(CInt(oRequest("ReportTitle").Item))
					sRowContents = sRowContents & RTF_FONT_END
					sRowContents = sRowContents & RTF_PARAGRAPH_END
					sRowContents = sRowContents & RTF_HEADER_END
					sRowContents = sRowContents & RTF_FOOTER_WITH_PAGE
					sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
					lErrorNumber = SaveTextToFile(sDocumentName, sRowContents, sErrorDescription)
				End If
				iRowP = 0
				If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
					mConcepts(iRowP,0) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					mConcepts(iRowP,1) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
					mConcepts(iRowP,2) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					mConcepts(iRowP,3) = FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
					mConcepts(iRowP,4) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					iRowP = iRowP + 1
					mConcepts(iRowP,0) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					mConcepts(iRowP,1) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
					mConcepts(iRowP,2) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					mConcepts(iRowP,3) = FormatNumber(CLng(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
					mConcepts(iRowP,4) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					iRowP = iRowP + 1
					iRowD = 0
					mConcepts(iRowD,5) = ""
					mConcepts(iRowD,6) = ""
					mConcepts(iRowD,7) = ""
					mConcepts(iRowD,8) = 0
					mConcepts(iRowD,9) = "0.00"
					iRowD = iRowD + 1
				ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
					mConcepts(iRowP,0) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					mConcepts(iRowP,1) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
					mConcepts(iRowP,2) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					mConcepts(iRowP,3) = FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
					mConcepts(iRowP,4) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					iRowP = iRowP + 1
					mConcepts(iRowP,0) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					mConcepts(iRowP,1) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
					mConcepts(iRowP,2) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					mConcepts(iRowP,3) = FormatNumber(CLng(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
					mConcepts(iRowP,4) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					iRowP = iRowP + 1
					iRowD = 0
					mConcepts(iRowD,5) = ""
					mConcepts(iRowD,6) = ""
					mConcepts(iRowD,7) = ""
					mConcepts(iRowD,8) = 0
					mConcepts(iRowD,9) = "0.00"
					iRowD = iRowD + 1
				Else
					Do While Not oRecordset.EOF
						mConcepts(iRowP,0) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
						mConcepts(iRowP,1) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
						mConcepts(iRowP,2) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
						mConcepts(iRowP,3) = FormatNumber(CLng(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
						mConcepts(iRowP,4) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
						iRowP = iRowP + 1
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
				End If
				oRecordset.Close

				sErrorDescription = "No se pudieron obtener las deducciones de los empleados registrados en el sistema."
				If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				Else
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				End If
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						iRowD = 0
						Do While Not oRecordset.EOF
							mConcepts(iRowD,5) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
							mConcepts(iRowD,6) = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value))
							mConcepts(iRowD,7) = CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
							mConcepts(iRowD,8) = FormatNumber(CLng(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
							mConcepts(iRowD,9) = FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
							iRowD = iRowD + 1
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
					End If
				End If

				iCount = 0
				bBorder = False
				asCellWidths = Split("400,1000,4000,5000,6400,7400,8000,11000,12000,13400", ",", -1, vbBinaryCompare)
				asCellAlignments = Split("RIGHT,CENTER,LEFT,RIGHT,RIGHT,RIGHT,CENTER,LEFT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

				sRowContents = RTF_PARAGRAPH_BEGIN & RTF_FONT15_START
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "CPTO;.;PTDA;.;PERCEPCIONES;.;EMPLEADOS;.;TOTAL;.;CPTO;.;PTDA;.;DEDUCCIONES;.;EMPLEADOS;.;TOTAL"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				'sRowContents = RTF_TABLE_BEGIN & RTF_ROW_BEGIN
				sRowContents = RTF_TABLE_BEGIN
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, bBorder, sDocumentName, sErrorDescription)
				sRowContents = RTF_TABLE_END
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										
				If iRowP > iRowD Then
					For iIndex = 1 To iRowP
						sErrorDescription = "No se desplegar la información del concentrado de pagos."
						If iRowD >= iCount Then
							sRowContents = mConcepts(iIndex,0)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,1)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,2)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,3)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,4)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,5)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,6)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,7)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,8)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,9)
						Else
							sRowContents = mConcepts(iIndex,0)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,1)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,2)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,3)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,4)
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
						End If
						iCount = iCount + 1
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						sRowContents = RTF_TABLE_BEGIN
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, bBorder, sDocumentName, sErrorDescription)
						sRowContents = RTF_TABLE_END
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					Next
				Else
					For iIndex = 1 To iRowD
						sErrorDescription = "No se desplegar la información del concentrado de pagos."
						If iRowP >= iCount Then
							sRowContents = mConcepts(iIndex,0)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,1)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,2)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,3)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,4)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,5)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,6)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,7)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,8)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,9)
						Else
							sRowContents = ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & ""
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,5)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,6)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,7)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,8)
							sRowContents = sRowContents & TABLE_SEPARATOR & mConcepts(iIndex,9)
						End If
						iCount = iCount + 1
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						sRowContents = RTF_TABLE_BEGIN
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, bBorder, sDocumentName, sErrorDescription)
						sRowContents = RTF_TABLE_END
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					Next
				End If

				sRowContents = RTF_FONT_END & RTF_PARAGRAPH_END
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

				sErrorDescription = "No se pudieron obtener las deducciones de los empleados registrados en el sistema."
				If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				Else
					If bPayrollIsClosed Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					End If
				End If
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sRowContents = RTF_PARAGRAPH_BEGIN
						sRowContents = sRowContents & RTF_FONT18_START
						sRowContents = sRowContents & RTF_TAB & " "
						sRowContents = sRowContents & "Total de empleados: "
							sRowContents = sRowContents & FormatNumber(CLng(oRecordset.Fields("TotalPayments").Value), 0, True, False, True)
						sRowContents = sRowContents & "   "
						sRowContents = sRowContents & "Total percepciones: " & mConcepts(0,4) & "   "
						sRowContents = sRowContents & "Total deducciones: " & mConcepts(0,9) & "   "
						sRowContents = sRowContents & "Total líquido: " & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & RFT_NEW_LINE

						sRowContents = sRowContents & RTF_FONT_END
						sRowContents = sRowContents & RTF_PARAGRAPH_END
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					oRecordset.Close
				End If
				sRowContents = RTF_END
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1006 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1007(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Remesa para cubrir la nómina de operativos. Reporte basado en la hoja 001125
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1007"
	Dim sContents
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim oRecordset
	Dim adTotal
	Dim asAccounts
	Dim iIndex
	Dim jIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Const S_STATE_IDS = "1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,19,20,-1,21,22,23,24,25,26,27,28,29,30,31,32,9"
	Dim asStateIDs
	Dim sStateName

	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sDate

	asStateIDs = Split(S_STATE_IDS, ",")
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "(Areas.", "(Areas1."), "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Concepts.", "Payroll_" & lPayrollID & "."), "EmployeeTypes.", "EmployeesHistoryList.")
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	If (InStr(1, sCondition, "Areas.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "Areas1.", vbBinaryCompare) > 0) Then
		'xxxxxxxxxxxxxxxxx
	End If

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1007.htm"), sErrorDescription)
	sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
	sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
	sContents = Replace(sContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
	Response.Write sContents
	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Delegación,Cuenta,Neto ISSSTE,Neto Vivienda,Neto Total", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
		asAccounts = ""
		sErrorDescription = "No se pudieron obtener las cuentas bancarias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StateID, AccountNumber From BankAccounts Where (StateID>0) And (EmployeeID=-1)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				asAccounts = asAccounts & CStr(oRecordset.Fields("StateID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AccountNumber").Value) & LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			If Len(asAccounts) > 0 Then asAccounts = Left(asAccounts, (Len(asAccounts) - Len(LIST_SEPARATOR)))
		End If
		asAccounts = Split(asAccounts, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asAccounts)
			asAccounts(iIndex) = Split(asAccounts(iIndex), SECOND_LIST_SEPARATOR)
			asAccounts(iIndex)(0) = CLng(asAccounts(iIndex)(0))
		Next
		adTotal = Split("0,0", ",")
		adTotal(0) = 0
		adTotal(1) = 0
		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "0", vbBinaryCompare) = 0) Then
			For iIndex = 0 To UBound(asStateIDs) - 1
				If CLng(asStateIDs(iIndex)) = -1 Then
					sStateName = "20A HOSP. REG. PDTE. JUAREZ OAXACA, OAX."
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID=38) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID=38) " & sCondition & " -->" & vbNewLine
						End If
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID=38) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID=38) " & sCondition & " -->" & vbNewLine
						End If
					Else
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (ConceptID=0) And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (ConceptID=0) And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (ConceptID=0) And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID=38) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (ConceptID=0) And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID=38) " & sCondition & " -->" & vbNewLine
						End If
					End If
				Else
					Call GetNameFromTable(oADODBConnection, "States", asStateIDs(iIndex), "", "", sStateName, "")
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
						End If
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
						End If
					Else
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
						End If
					End If
				End If
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sRowContents = CleanStringForHTML(sStateName)
						If CLng(asStateIDs(iIndex)) = -1 Then
							For jIndex = 0 To UBound(asAccounts)
								If asAccounts(jIndex)(0) = 20 Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(asAccounts(jIndex)(1))
									Exit For
								End If
							Next
						Else
							For jIndex = 0 To UBound(asAccounts)
								If asAccounts(jIndex)(0) = CLng(asStateIDs(iIndex)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(asAccounts(jIndex)(1))
									Exit For
								End If
							Next
						End If
						If IsNull(oRecordset.Fields("TotalAmount")) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
							sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
							sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						End If

						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
				End If
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
			Next

			sRowContents = "<SPAN COLS=""2"" /><B>TOTAL FORÁNEOS</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0.00</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If

		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "1", vbBinaryCompare) = 0) Then
			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			Call GetNameFromTable(oADODBConnection, "States", 9, "", "", sStateName, "")
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=124) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=155) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (ConceptID=0) And (Areas1.AreaID<>38) And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " -->" & vbNewLine
				End If
			End If
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(sStateName)
					For jIndex = 0 To UBound(asAccounts)
						If asAccounts(jIndex)(0) = 9 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(asAccounts(jIndex)(1))
							Exit For
						End If
					Next
					If IsNull(oRecordset.Fields("TotalAmount")) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
						sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
						sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
						adTotal(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
						sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
						adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If

			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "<SPAN COLS=""2"" /><B>TOTAL LOCAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0.00</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If

		If (Len(oRequest("StateType").Item) = 0) Then
			sRowContents = "<SPAN COLS=""2"" /><B>TOTAL GENERAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0.00</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If
	Response.Write "</TABLE>"

	Set oRecordset = Nothing
	BuildReport1007 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1008(oRequest, oADODBConnection, bShowTotals, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte por empresa y por tipo de empleado. Reporte basado en la hoja 001129
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bShowTotals, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1008"
	
	Set oRecordset = Nothing
	BuildReport1008 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1009(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Resumen mensual de nóminas. Reporte basado en la hoja 001138
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1009"
	Dim sContents
	Dim sCondition
	Dim sCondition2
	Dim lForPayrollID
	Dim asCLCs
	Dim adTotals
	Dim asParameters
	Dim iIndex
	Dim jIndex
	Dim iCounter
	Dim sCurrentID
	Dim sNames
	Dim oRecordset
	Dim sRowContents
	Dim asColumnsTitles
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "Companies.", "EmployeesHistoryListForPayroll."), "Employees.", "EmployeesHistoryListForPayroll."), "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "EmployeeTypes.", "EmployeesHistoryListForPayroll."), "PaymentCenters.AreaID", "EmployeesHistoryListForPayroll.PaymentCenterID")
	lForPayrollID = (CLng(oRequest("YearID").Item) * 100) + CInt(oRequest("MonthID").Item)

	sErrorDescription = "No se pudieron obtener el resumen mensual de nóminas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, PayrollsCLCs.PayrollCLC, PayrollsCLCs.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID, Count(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalCount, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payrolls, Payroll_" & oRequest("YearID").Item & ", PayrollsCLCs, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (Payrolls.PayrollID=Payroll_" & oRequest("YearID").Item & ".RecordDate) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID=PayrollsCLCs.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".RecordDate=PayrollsCLCs.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & oRequest("YearID").Item & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID<>0) And (EmployeesHistoryListForPayroll.PayrollID>=" & lForPayrollID & "00) And (EmployeesHistoryListForPayroll.PayrollID<=" & lForPayrollID & "99) And (Areas.StartDate<=" & lForPayrollID & "00) And (Areas.EndDate>=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID In (0,-1,-2)) " & sCondition & " Group By Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, PayrollsCLCs.PayrollCLC, PayrollsCLCs.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By Payrolls.PayrollID, PayrollsCLCs.PayrollCLC, Payroll_" & oRequest("YearID").Item & ".ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, PayrollsCLCs.PayrollCLC, PayrollsCLCs.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID, Count(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalCount, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payrolls, Payroll_" & oRequest("YearID").Item & ", PayrollsCLCs, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (Payrolls.PayrollID=Payroll_" & oRequest("YearID").Item & ".RecordDate) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID=PayrollsCLCs.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".RecordDate=PayrollsCLCs.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & oRequest("YearID").Item & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID<>0) And (EmployeesHistoryListForPayroll.PayrollID>=" & lForPayrollID & "00) And (EmployeesHistoryListForPayroll.PayrollID<=" & lForPayrollID & "99) And (Areas.StartDate<=" & lForPayrollID & "00) And (Areas.EndDate>=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID In (0,-1,-2)) " & sCondition & " Group By Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, PayrollsCLCs.PayrollCLC, PayrollsCLCs.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By Payrolls.PayrollID, PayrollsCLCs.PayrollCLC, Payroll_" & oRequest("YearID").Item & ".ConceptID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1007.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
			sContents = Replace(sContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
			Response.Write sContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No.,CLC/LOCAL,UNIDAD,QNA.,BANCO,F/PAGO,AREA,ARCHIVO,REG.,PERCEPCIONES,DEDUCCIONES,LIQUIDO", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,,,,,,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				iCounter = 1
				asCLCs = ""
				adTotals = Split(",", ",")
				adTotals(0) = Split("0,0,0,0", ",")
				adTotals(1) = Split("0,0,0,0", ",")
				For iIndex = 0 To UBound(adTotals(0))
					adTotals(0)(iIndex) = 0
					adTotals(1)(iIndex) = 0
				Next
				sCurrentID = ""
				Do While Not oRecordset.EOF
					If StrComp(sCurrentID, CStr(oRecordset.Fields("PayrollID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PayrollCLC").Value), vbBinaryCompare) <> 0 Then
						If Len(sCurrentID) > 0 Then
							asCLCs = Replace(asCLCs, "<TOTAL_COUNT />", "0.00")
							asCLCs = Replace(asCLCs, "<CONCEPT_1 />", "0.00")
							asCLCs = Replace(asCLCs, "<CONCEPT_2 />", "0.00")
							asCLCs = Replace(asCLCs, "<CONCEPT_0 />", "0.00")
							asCLCs = asCLCs & LIST_SEPARATOR
							iCounter = iCounter + 1
						End If
						asCLCs = asCLCs & iCounter '(0)
						asCLCs = asCLCs & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PayrollCLC").Value)) '(1)
						If CLng(oRecordset.Fields("PayrollTypeID").Value) <> 1 Then '(2)
							asCLCs = asCLCs & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PayrollName").Value))
						Else
							asCLCs = asCLCs & TABLE_SEPARATOR & "Ordinaria"
						End If
						asCLCs = asCLCs & TABLE_SEPARATOR & "---" 'CompanyName (3)
						'asCLCs = asCLCs & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1) 'QNA (4)
						asCLCs = asCLCs & TABLE_SEPARATOR & GetPayrollNumber(CLng(oRecordset.Fields("PayrollID").Value)) & "/" & Left(CLng(oRecordset.Fields("PayrollID").Value), Len("0000")) 'QNA (4)
						asCLCs = asCLCs & TABLE_SEPARATOR & "---" 'BankName (5)
						asCLCs = asCLCs & TABLE_SEPARATOR & "---" 'PaymentType (6)
						asCLCs = asCLCs & TABLE_SEPARATOR & "---" 'AreaType (7)
						asCLCs = asCLCs & TABLE_SEPARATOR & "---" 'FileName (8)
						asCLCs = asCLCs & TABLE_SEPARATOR & CStr(oRecordset.Fields("FilterParameters").Value) '(9)
						asCLCs = asCLCs & TABLE_SEPARATOR & "<TOTAL_COUNT />" '(10)
						asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_1 />" '(11)
						asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_2 />" '(12)
						asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_0 />" '(13)
						sCurrentID = CStr(oRecordset.Fields("PayrollID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PayrollCLC").Value)
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case -2
							asCLCs = Replace(asCLCs, "<CONCEPT_2 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							adTotals(0)(2) = adTotals(0)(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case -1
							asCLCs = Replace(asCLCs, "<CONCEPT_1 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							adTotals(0)(1) = adTotals(0)(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 0
							asCLCs = Replace(asCLCs, "<CONCEPT_0 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							asCLCs = Replace(asCLCs, "<TOTAL_COUNT />", FormatNumber(CLng(oRecordset.Fields("TotalCount").Value), 0, True, False, True))
							adTotals(0)(0) = adTotals(0)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adTotals(0)(3) = adTotals(0)(3) + CLng(oRecordset.Fields("TotalCount").Value)
					End Select
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				asCLCs = Replace(asCLCs, "<TOTAL_COUNT />", "0.00")
				asCLCs = Replace(asCLCs, "<CONCEPT_1 />", "0.00")
				asCLCs = Replace(asCLCs, "<CONCEPT_2 />", "0.00")
				asCLCs = Replace(asCLCs, "<CONCEPT_0 />", "0.00")
				asCLCs = Split(asCLCs, LIST_SEPARATOR)
				For iIndex = 0 To UBound(asCLCs)
					asCLCs(iIndex) = Split(asCLCs(iIndex), TABLE_SEPARATOR)
					sNames = ""
					asCLCs(iIndex)(8) = "<EMPLOYEE_TYPES /><AREAS /><PAYMENT_CENTERS /><EMPLOYEES />"
					asParameters = Split(asCLCs(iIndex)(9), " And ")
					For jIndex = 0 To UBound(asParameters)
						If InStr(1, asParameters(jIndex), "EmployeesHistoryList.EmployeeID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID In (", ""), "))", "")
							asCLCs(iIndex)(8) = Replace(asCLCs(iIndex)(8), "<EMPLOYEES />", sCondition2 & "&nbsp;")
						End If

						If InStr(1, asParameters(jIndex), "EmployeesHistoryList.CompanyID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryList.CompanyID In (", ""), "))", "")
							Call GetNameFromTable(oADODBConnection, "Companies", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							If Len(sCondition2) > 0 Then asCLCs(iIndex)(3) = sCondition2
						End If

						If InStr(1, asParameters(jIndex), "EmployeesHistoryList.EmployeeTypeID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryList.EmployeeTypeID In (", ""), "))", "")
							Call GetNameFromTable(oADODBConnection, "EmployeeTypes", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							asCLCs(iIndex)(8) = Replace(asCLCs(iIndex)(8), "<EMPLOYEE_TYPES />", sCondition2 & "&nbsp;")
						End If

						If InStr(1, asParameters(jIndex), "Areas.AreaID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaID In (", ""), "))", "")
							Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							asCLCs(iIndex)(8) = Replace(asCLCs(iIndex)(8), "<AREAS />", sCondition2 & "&nbsp;")
						End If
						If InStr(1, asParameters(jIndex), "Areas.AreaPath ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaPath Like ""%,", ""), ",%"")", "")
							Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							asCLCs(iIndex)(8) = Replace(asCLCs(iIndex)(8), "<AREAS />", sCondition2 & "&nbsp;")
						End If

						If InStr(1, asParameters(jIndex), "EmployeesHistoryList.PaymentCenterID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryList.PaymentCenterID In (", ""), "))", "")
							Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							asCLCs(iIndex)(8) = Replace(asCLCs(iIndex)(8), "<PAYMENT_CENTERS />", sCondition2 & "&nbsp;")
						End If

						If InStr(1, asParameters(jIndex), "ParentZones.ZoneID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(ParentZones.ZoneID In (", ""), "))", "")
							If StrComp(sCondition2, "9", vbBinaryCompare) = 0 Then
								sCondition2 = "Local"
							Else
								Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							End If
							If Len(sCondition2) > 0 Then asCLCs(iIndex)(3) = sCondition2
						End If
						If InStr(1, asParameters(jIndex), "Zones.ZonePath ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(Zones.ZonePath Like ""%,", ""), ",%"")", "")
							If StrComp(sCondition2, "9", vbBinaryCompare) = 0 Then
								sCondition2 = "Local"
							Else
								Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							End If
							If Len(sCondition2) > 0 Then asCLCs(iIndex)(3) = sCondition2
						End If

						If InStr(1, asParameters(jIndex), "BankAccounts.BankID ", vbBinaryCompare) > 0 Then
							sCondition2 = Replace(Replace(asParameters(jIndex), "(BankAccounts.BankID In (", ""), "))", "")
							Call GetNameFromTable(oADODBConnection, "Banks", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
							If Len(sCondition2) > 0 Then asCLCs(iIndex)(5) = sCondition2
						End If

						If StrComp(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID<600000)", vbBinaryCompare) = 0 Then
							If UBound(asParameters) <= jIndex Then
								sCondition2 = "Empleados con cheque y depósitos"
							Else
								If (StrComp(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber<>""."")", vbBinaryCompare) = 0) Then
									sCondition2 = "Empleados con depósito"
								ElseIf (StrComp(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber=""."")", vbBinaryCompare) = 0) Then
									sCondition2 = "Empleados con cheque"
								End If
							End If
						ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID>=600000)", vbBinaryCompare) = 0 Then
							sCondition2 = "Honorarios"
						ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryList.EmployeeID>=700000)", vbBinaryCompare) = 0 Then
							sCondition2 = "Pensión alimenticia"
						End If
						If Len(sCondition2) > 0 Then asCLCs(iIndex)(3) = sCondition2
					Next
					asCLCs(iIndex)(8) = Replace(Replace(Replace(Replace(asCLCs(iIndex)(8), "<EMPLOYEE_TYPES />", ""), "<AREAS />", ""), "<PAYMENT_CENTERS />", ""), "<EMPLOYEES />", "")

					sRowContents = asCLCs(iIndex)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(asCLCs(iIndex)(1))
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(3)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(4)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(5)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(6)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(7)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(8)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(10)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(11)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(12)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(13)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				Next
				sRowContents = "<SPAN COLS=""8"" /><B>TOTAL" & TABLE_SEPARATOR & FormatNumber(adTotals(0)(3), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(1), 2, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(2), 2, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(0), 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = -1
			sErrorDescription = "No existen CLCs generadas que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1009 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1010(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Ramas médica, paramédica, de grupos afines y operativa, de enlace y enlacede alto nivel de responsabilidad. Reporte basado en la hoja 001144
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1010"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sPeriod
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim oRecordset
	Dim asZones
	Dim sRowContents
	Dim asRowContents
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lCounter1
	Dim lCounter2
	Dim dPerceptions1
	Dim dPerceptions2
	Dim dDeductions1
	Dim dDeductions2
	Dim dTotal1
	Dim dTotal2

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "PaymentCenters.AreaID", "EmployeesHistoryList.PaymentCenterID")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & sDate & ".doc"				

	lCounter1 = 0
	lCounter2 = 0
	dPerceptions1 = 0
	dPerceptions2 = 0
	dDeductions1 = 0
	dDeductions2 = 0
	dTotal1 = 0
	dTotal2 = 0
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sRowContents = RTF_BEGIN_V & " "
		sRowContents = sRowContents & RTF_DEFAULT_TITLE
		sRowContents = sRowContents & RTF_HEADER_BEGIN
		sRowContents = sRowContents & RTF_FONT18_START
		sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN
		lErrorNumber = SaveTextToFile(sDocumentName, sRowContents, sErrorDescription)
		sRowContents = "{ " & GetFileContents(Server.MapPath("Templates\LogoISSSTE_RTF.txt"), sErrorDescription)
		sRowContents = sRowContents & TABLE_SEPARATOR
		sRowContents = sRowContents & aReportTitle(L_COMPANY_FLAGS) & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "OPERATIVOS" & RFT_NEW_LINE & RFT_NEW_LINE & " "
		If Len(aReportTitle(L_BANK_FLAGS)) > 0 Then
			sRowContents = sRowContents & UCase(aReportTitle(L_BANK_FLAGS)) & RFT_NEW_LINE & " "
		End If
		sRowContents = sRowContents & UCase(aReportTitle(L_PAYROLL_FLAGS)) & RFT_NEW_LINE & " "
		sRowContents = sRowContents & TABLE_SEPARATOR
		sRowContents = sRowContents & RTF_PAGE_NUMBER & " " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "FECHA: " & DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) & " " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "HORA: " & DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))) & " " & RFT_NEW_LINE & " "
		asCellAlignments = Split("LEFT,CENTER,RIGHT", ",", -1, vbBinaryCompare)
		asCellWidths = Split("3000,8000,10000", ",", -1, vbBinaryCompare)
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, False, sDocumentName, sErrorDescription)
		sRowContents = "RAMAS MÉDICA, PARAMÉDICA, DE GRUPOS AFINES Y OPERATIVA, DE ENLACE Y " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "ENLACE DE ALTO NIVEL DE RESPONSABILIDAD " & RFT_NEW_LINE & " "
		asCellAlignments = Split("CENTER", ",", -1, vbBinaryCompare)
		asCellWidths = Split("10000", ",", -1, vbBinaryCompare)
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, False, sDocumentName, sErrorDescription)
		sRowContents = "}"
		sRowContents = sRowContents & RTF_PARAGRAPH_END
		sRowContents = sRowContents & RTF_FONT_END
		sRowContents = sRowContents & RTF_HEADER_END
		sRowContents = sRowContents & RTF_FOOTER_BEGIN & " "
		sRowContents = sRowContents & RTF_FOOTER_END & " "
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName From Zones Where (Zones.ParentID=-1) And (Zones.ZoneID>0) Order By ZoneCode", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asZones = ""
			Do While Not oRecordset.EOF
				asZones = asZones & oRecordset.Fields("ZoneID").Value & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ZoneCode").Value) & " " & CStr(oRecordset.Fields("ZoneName").Value) & SECOND_LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR & "0" & LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asZones = Left(asZones, (Len(asZones) - Len(LIST_SEPARATOR)))
			asZones = Split(asZones, LIST_SEPARATOR)

			sRowContents = RTF_PARAGRAPH_BEGIN & RTF_FONT15_START
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "FORANEO;.;REGISTROS;.;PERCEPCIONES;.;DEDUCCIONES;.;LIQUIDO"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellWidths = Split("4000,5500,7000,8500,10000", ",", -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			asCellAlignments = Split("LEFT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
			For iIndex = 0 To UBound(asZones)
				asZones(iIndex) = Split(asZones(iIndex), SECOND_LIST_SEPARATOR)
				sErrorDescription = "No se pudo obtener la información de la nómina"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As TotalCounter, Sum(ConceptAmount) As TotalPayments, ConceptID From BankAccounts, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryList.EmployeeTypeID In (0,2,3,4,5)) And (Payroll_" & lPayrollID & ".ConceptID In (-2,-1,0)) And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & asZones(iIndex)(0) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						If Not IsNull(oRecordset.Fields("TotalCounter").Value) Then asZones(iIndex)(2) = CLng(oRecordset.Fields("TotalCounter").Value)
						Do While Not oRecordset.EOF
							If Not IsNull(oRecordset.Fields("TotalPayments").Value) Then asZones(iIndex)(CInt(oRecordset.Fields("ConceptID").Value) + 5) = CDbl(oRecordset.Fields("TotalPayments").Value)
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
					End If
					oRecordset.Close
				End If
			Next
			For iIndex = 0 To UBound(asZones)
				sRowContents = asZones(iIndex)(1)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asZones(iIndex)(2), 0, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asZones(iIndex)(4), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asZones(iIndex)(3), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asZones(iIndex)(5), 2, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
				If CLng(asZones(iIndex)(0)) <> 9 Then
					lCounter1 = lCounter1 + CLng(asZones(iIndex)(2))
					dPerceptions1 = dPerceptions1 + CDbl(asZones(iIndex)(4))
					dDeductions1 = dDeductions1 + CDbl(asZones(iIndex)(3))
					dTotal1 = dTotal1 + CDbl(asZones(iIndex)(5))
				Else
					lCounter2 = CLng(asZones(iIndex)(2))
					dPerceptions2 = CDbl(asZones(iIndex)(4))
					dDeductions2 = CDbl(asZones(iIndex)(3))
					dTotal2 = CDbl(asZones(iIndex)(5))
				End If
			Next
			sRowContents = "TOTAL FORANEO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter1, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions1, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dDeductions1, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal1, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			sRowContents = "LOCAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter2, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dDeductions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal2, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			sRowContents = "TOTAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter1 + lCounter2, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions1 + dPerceptions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dDeductions1 + dDeductions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal1 + dTotal2, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			sRowContents = RTF_FONT_END & RTF_PARAGRAPH_END
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = RFT_BRAKE_PAGE
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		End If

		sRowContents = RTF_PARAGRAPH_BEGIN & RTF_FONT15_START
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		lErrorNumber = BuildReport1010_Resume(oRequest, oADODBConnection, bSpecifiedZone, sDocumentName, sErrorDescription)
		sRowContents = RTF_FONT_END & RTF_PARAGRAPH_END
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

		sRowContents = RTF_END
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudieron guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
		End If
		oEndDate = Now()
		If (lErrorNumber = 0) And B_USE_SMTP Then
			If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1010 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1010_Resume(oRequest, oADODBConnection, bSpecifiedZone, sDocumentName, sErrorDescription)
'************************************************************
'Purpose: Ramas médica, paramédica, de grupos afines y operativa, de enlace y enlacede alto nivel de responsabilidad. Reporte basado en la hoja 001144
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bSpecifiedZone, sDocumentName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1010_Resume"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sPeriod
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim oRecordset
	Dim sRowContents
	Dim asRowContents
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim Total
	Dim TotalPayments
	Dim TotalDeductions
	Dim TotalPaid

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	If (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Companies.", "Employees.")
	End If
	If (InStr(1, sCondition, "EmployeeTypes.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "EmployeeTypes.", "Employees.")
	End If
	If (InStr(1, sCondition, "Banks.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Banks.", "BankAccounts.")
	End If

	If Not bSpecifiedZone Then
		sCondition = sCondition & " And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "')"
	End If

	sDate = GetSerialNumberForDate("")
	sErrorDescription = "No se pudo obtener la información del empleado."

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(Percepciones.ConceptAmount) As TotalPayments, Sum(Deducciones.ConceptAmount) As TotalDeductions From BankAccounts, Employees, Jobs, Areas, Zones, Payroll_" & lPayrollID & " As Percepciones, Payroll_" & lPayrollID & " As Deducciones Where (BankAccounts.EmployeeID = Employees.EmployeeID) And (Percepciones.EmployeeID = BankAccounts.EmployeeID) And (Percepciones.ConceptID = -1) And (Deducciones.EmployeeID = BankAccounts.EmployeeID) And (Deducciones.ConceptID = -2) And (EmployeeTypeID In (0,2,3,4,5)) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sRowContents = "DESCRIPCIÓN;.;REGISTROS;.;PERCEPCIONES;.;DEDUCCIONES;.;LIQUIDO"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellWidths = Split("4000,5500,7000,8500,10000", ",", -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			If bSpecifiedZone Then
				If (InStr(1, aReportTitle(L_ZONE_FLAGS), "DISTRITO FEDERAL", vbBinaryCompare) > 0) Then
					aReportTitle(L_ZONE_FLAGS) = Replace(aReportTitle(L_ZONE_FLAGS), "DISTRITO FEDERAL", "LOCAL")
				End If
				sRowContents = "DÉBITO " & aReportTitle(L_ZONE_FLAGS)
			Else
				sRowContents = "DÉBITO LOCAL"
			End If
				
			Total = CDbl(oRecordset.Fields("Total").Value)
			TotalPayments = CDbl(oRecordset.Fields("TotalPayments").Value)
			TotalDeductions = CDbl(oRecordset.Fields("TotalDeductions").Value)
			TotalPaid = TotalPayments - TotalDeductions
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(Total, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPayments, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalDeductions, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPaid, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			If Not bSpecifiedZone Then
				oRecordset.Close
				sCondition = Replace(sCondition, "Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "'", "Zones.ZonePath Not Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "'")
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(Percepciones.ConceptAmount) As TotalPayments, Sum(Deducciones.ConceptAmount) As TotalDeductions From BankAccounts, Employees, Jobs, Areas, Zones, Payroll_" & lPayrollID & " As Percepciones, Payroll_" & lPayrollID & " As Deducciones Where (Employees.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.ConceptID=-1) And (Deducciones.EmployeeID=BankAccounts.EmployeeID) And (Deducciones.ConceptID=-2) And (EmployeeTypeID In (0,2,3,4,5)) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sRowContents = "DEBITO FORÁNEO"
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Total").Value), 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalDeductions").Value), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((CDbl(oRecordset.Fields("TotalPayments").Value)-CDbl(oRecordset.Fields("TotalDeductions").Value)), 2, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
						Total = Total + CDbl(oRecordset.Fields("Total").Value)
						TotalPayments = TotalPayments + CDbl(oRecordset.Fields("TotalPayments").Value)
						TotalDeductions = TotalDeductions + CDbl(oRecordset.Fields("TotalDeductions").Value)
						TotalPaid = TotalPaid + TotalPayments - TotalDeductions
					End If
				End If
			End If
			sRowContents = "TOTAL"
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(Total, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPayments, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalDeductions, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPaid, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1010_Resume = lErrorNumber
	Err.Clear
End Function

Function BuildReport1011(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Pensión alimenticia de Ramas médica, paramédica, de grupos afines y operativa, de enlace y enlacede alto nivel de responsabilidad. Reporte basado en la hoja 001164
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1011"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sPeriod
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim oRecordset
	Dim asZones
	Dim sRowContents
	Dim asRowContents
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lCounter1
	Dim lCounter2
	Dim dPerceptions1
	Dim dPerceptions2
	Dim dTotal1
	Dim dTotal2

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "PaymentCenters.AreaID", "EmployeesHistoryList.PaymentCenterID")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & sDate & ".doc"				

	lCounter1 = 0
	lCounter2 = 0
	dPerceptions1 = 0
	dPerceptions2 = 0
	dTotal1 = 0
	dTotal2 = 0
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sRowContents = RTF_BEGIN_V & " "
		sRowContents = sRowContents & RTF_DEFAULT_TITLE
		sRowContents = sRowContents & RTF_HEADER_BEGIN
		sRowContents = sRowContents & RTF_FONT18_START
		sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN
		lErrorNumber = SaveTextToFile(sDocumentName, sRowContents, sErrorDescription)
		sRowContents = "{ " & GetFileContents(Server.MapPath("Templates\LogoISSSTE_RTF.txt"), sErrorDescription)
		sRowContents = sRowContents & TABLE_SEPARATOR
		sRowContents = sRowContents & aReportTitle(L_COMPANY_FLAGS) & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "OPERATIVOS" & RFT_NEW_LINE & RFT_NEW_LINE & " "
		If Len(aReportTitle(L_BANK_FLAGS)) > 0 Then
			sRowContents = sRowContents & UCase(aReportTitle(L_BANK_FLAGS)) & RFT_NEW_LINE & " "
		End If
		sRowContents = sRowContents & "PENSIÓN ALIMENTICIA" & RFT_NEW_LINE & " "
		sRowContents = sRowContents & UCase(aReportTitle(L_PAYROLL_FLAGS)) & RFT_NEW_LINE & " "
		sRowContents = sRowContents & TABLE_SEPARATOR
		sRowContents = sRowContents & RTF_PAGE_NUMBER & " " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "FECHA: " & DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) & " " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "HORA: " & DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))) & " " & RFT_NEW_LINE & " "
		asCellAlignments = Split("LEFT,CENTER,RIGHT", ",", -1, vbBinaryCompare)
		asCellWidths = Split("3000,8000,10000", ",", -1, vbBinaryCompare)
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, False, sDocumentName, sErrorDescription)
		sRowContents = "RAMAS MÉDICA, PARAMÉDICA, DE GRUPOS AFINES Y OPERATIVA, DE ENLACE Y " & RFT_NEW_LINE & " "
		sRowContents = sRowContents & "ENLACE DE ALTO NIVEL DE RESPONSABILIDAD " & RFT_NEW_LINE & " "
		asCellAlignments = Split("CENTER", ",", -1, vbBinaryCompare)
		asCellWidths = Split("10000", ",", -1, vbBinaryCompare)
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, False, sDocumentName, sErrorDescription)
		sRowContents = "}"
		sRowContents = sRowContents & RTF_PARAGRAPH_END
		sRowContents = sRowContents & RTF_FONT_END
		sRowContents = sRowContents & RTF_HEADER_END
		sRowContents = sRowContents & RTF_FOOTER_BEGIN & " "
		sRowContents = sRowContents & RTF_FOOTER_END & " "
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName From Zones Where (Zones.ParentID=-1) And (Zones.ZoneID>0) Order By ZoneCode", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asZones = ""
			Do While Not oRecordset.EOF
				asZones = asZones & oRecordset.Fields("ZoneID").Value & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ZoneCode").Value) & " " & CStr(oRecordset.Fields("ZoneName").Value) & LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asZones = Left(asZones, (Len(asZones) - Len(LIST_SEPARATOR)))
			asZones = Split(asZones, LIST_SEPARATOR)

			sRowContents = RTF_PARAGRAPH_BEGIN & RTF_FONT15_START
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "FORANEO;.;REGISTROS;.;PERCEPCIONES;.;DEDUCCIONES;.;LIQUIDO"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellWidths = Split("4000,5500,7000,8500,10000", ",", -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			asCellAlignments = Split("LEFT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
			For iIndex = 0 To UBound(asZones)
				asZones(iIndex) = Split(asZones(iIndex), SECOND_LIST_SEPARATOR)
				sErrorDescription = "No se pudo obtener la información de la nómina"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalPayments From BankAccounts, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=70) And (EmployeesHistoryList.EmployeeTypeID In (0,2,3,4,5)) " & sCondition & " And (ParentZones.ZoneID=" & asZones(iIndex)(0) & ")", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						If CLng(asZones(iIndex)(0)) <> 9 Then
							sRowContents = asZones(iIndex)(1)
							If Not IsNull(oRecordset.Fields("TotalPayments").Value) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Total").Value), 0, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 2, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 2, True, False, True)
								lCounter1 = lCounter1 + CDbl(oRecordset.Fields("Total").Value)
								dPerceptions1 = dPerceptions1 + CDbl(oRecordset.Fields("TotalPayments").Value)
								dTotal1 = dTotal1 + dPerceptions2
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "0" & TABLE_SEPARATOR & "0.00" & TABLE_SEPARATOR & "0.00" & TABLE_SEPARATOR & "0.00"
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
						Else
							If Not IsNull(oRecordset.Fields("TotalPayments").Value) Then
								lCounter2 = CDbl(oRecordset.Fields("Total").Value)
								dPerceptions2 = CDbl(oRecordset.Fields("TotalPayments").Value)
								dTotal2 = dPerceptions2
							End If
						End If
					End If
				End If
				oRecordset.Close
				If Err.number <> 0 Then Exit For
			Next
			sRowContents = "TOTAL FORANEO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter1, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions1, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal1, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			sRowContents = "LOCAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter2, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal2, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			sRowContents = "TOTAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(lCounter1 + lCounter2, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dPerceptions1 + dPerceptions2, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal1 + dTotal2, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			sRowContents = RTF_FONT_END & RTF_PARAGRAPH_END
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = RFT_BRAKE_PAGE
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		End If

		sRowContents = RTF_PARAGRAPH_BEGIN & RTF_FONT15_START
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		lErrorNumber = BuildReport1010_PA_Resume(oRequest, oADODBConnection, bSpecifiedZone, sDocumentName, sErrorDescription)
		sRowContents = RTF_FONT_END & RTF_PARAGRAPH_END
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		sRowContents = RTF_END
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudieron guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
		End If
		oEndDate = Now()
		If (lErrorNumber = 0) And B_USE_SMTP Then
			If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1011 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1011_Resume(oRequest, oADODBConnection, bSpecifiedZone, sDocumentName, sErrorDescription)
'************************************************************
'Purpose: Pensión alimenticia de ramas médica, paramédica, de grupos afines y operativa, de enlace y enlacede alto nivel de responsabilidad. Reporte basado en la hoja 001164
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bSpecifiedZone, sDocumentName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1011_Resume"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sPeriod
	Dim sDate
	Dim iIndex
	Dim jIndex
	Dim oRecordset
	Dim sRowContents
	Dim asRowContents
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim Total
	Dim TotalPayments
	Dim TotalDeductions
	Dim TotalPaid

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	If (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Companies.", "Employees.")
	End If
	If (InStr(1, sCondition, "EmployeeTypes.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "EmployeeTypes.", "Employees.")
	End If
	If (InStr(1, sCondition, "Banks.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Banks.", "BankAccounts.")
	End If

	If Not bSpecifiedZone Then
		sCondition = sCondition & " And (Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "')"
	End If

	sDate = GetSerialNumberForDate("")
	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(Percepciones.ConceptAmount) As TotalPayments, 0 As TotalDeductions From BankAccounts, Employees, Jobs, Areas, Zones, Payroll_" & lPayrollID & " As Percepciones Where (Employees.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.ConceptID=70) And (EmployeeTypeID In (0,2,3,4,5)) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sRowContents = "DESCRIPCIÓN;.;REGISTROS;.;PERCEPCIONES;.;DEDUCCIONES;.;LIQUIDO"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellWidths = Split("4000,5500,7000,8500,10000", ",", -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
			If bSpecifiedZone Then
				If (InStr(1, aReportTitle(L_ZONE_FLAGS), "DISTRITO FEDERAL", vbBinaryCompare) > 0) Then
					aReportTitle(L_ZONE_FLAGS) = Replace(aReportTitle(L_ZONE_FLAGS), "DISTRITO FEDERAL", "LOCAL")
				End If
				sRowContents = "DÉBITO " & aReportTitle(L_ZONE_FLAGS)
			Else
				sRowContents = "DÉBITO LOCAL"
			End If
				
			Total = CDbl(oRecordset.Fields("Total").Value)
			TotalPayments = CDbl(oRecordset.Fields("TotalPayments").Value)
			TotalDeductions = CDbl(oRecordset.Fields("TotalDeductions").Value)
			TotalPaid = TotalPayments - TotalDeductions
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(Total, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPayments, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalDeductions, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPaid, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			asCellAlignments = Split("CENTER,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

			If Not bSpecifiedZone Then
				oRecordset.Close
				sCondition = Replace(sCondition, "Zones.ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "'", "Zones.ZonePath Not Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "'")
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(Percepciones.ConceptAmount) As TotalPayments, 0 As TotalDeductions From BankAccounts, Employees, Jobs, Areas, Zones, Payroll_" & lPayrollID & " As Percepciones Where (Employees.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.EmployeeID=BankAccounts.EmployeeID) And (Percepciones.ConceptID=70) And (EmployeeTypeID In (0,2,3,4,5)) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sRowContents = "DEBITO FORÁNEO"
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Total").Value), 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalPayments").Value), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalDeductions").Value), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((CDbl(oRecordset.Fields("TotalPayments").Value)-CDbl(oRecordset.Fields("TotalDeductions").Value)), 2, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
						Total = Total + CDbl(oRecordset.Fields("Total").Value)
						TotalPayments = TotalPayments + CDbl(oRecordset.Fields("TotalPayments").Value)
						TotalDeductions = TotalDeductions + CDbl(oRecordset.Fields("TotalDeductions").Value)
						TotalPaid = TotalPaid + TotalPayments - TotalDeductions
					End If
				End If
			End If
			sRowContents = "TOTAL"
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(Total, 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPayments, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalDeductions, 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & FormatNumber(TotalPaid, 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)

		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1011_Resume = lErrorNumber
	Err.Clear
End Function

Function BuildReport1012(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Diferencias totales por concepto. Reporte basado en la hoja 001171
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1012"
	Dim sHeaderContents
	Dim iMonth
	Dim iYear
	Dim iIndex
	Dim asPayrollIDs
	Dim lPayments
	Dim dPayments
	Dim sQuery
	Dim oRecordset
	Dim oRecordset1
	Dim oRecordset2
	Dim asColumnsTitles
	Dim asColumnsTitles2
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iConceptID
	Dim dTotalAmount1
	Dim dTotalPayments1
	Dim dTotalAmount2
	Dim dTotalPayments2
	Dim iCount

	
	iMonth = CInt(Mid(oRequest.Item("PayrollID").Item, 5, 2))
	iYear = CInt(Mid(oRequest.Item("PayrollID").Item, 1, 4))
	If iMonth = 12 Then
		iMonth = 1
		iYear = iYear + 1
	Else
		iMonth = iMonth + 1
	End If

	sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollDate>=" & oRequest.Item("PayrollID").Item & ") And (PayrollDate<=" & (iYear & Right(("0" & iMonth), Len("00"))) & "99) And (PayrollTypeID = 1) Order By PayrollID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrollIDs = ""
	iCount = 0
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			iCount = iCount + 1
			asPayrollIDs = asPayrollIDs & CStr(oRecordset.Fields("PayrollID").Value) & ";"
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	oRecordset.Close
	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1012.htm"), sErrorDescription)
	If (Len(asPayrollIDs) > 0) Then
		asPayrollIDs = Left(asPayrollIDs, (Len(asPayrollIDs) - Len(";")))
		asPayrollIDs = Split(asPayrollIDs, ";")
	End If	
	
	If (Len(sHeaderContents) > 0) Then
		sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
		sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE1 />", DisplayNumericDateFromSerialNumber(asPayrollIDs(0)))
		sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE2 />", DisplayNumericDateFromSerialNumber(asPayrollIDs(1)))
		Response.Write sHeaderContents
		
		sErrorDescription = "No se pudieron obtener los conceptos de pago."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptShortName, ConceptName, IsDeduction From Concepts Where (ConceptID>0) And (IsDeduction In(0,1)) And (EndDate=30000000) Order by ConceptShortName", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asColumnsTitles2 = Split("&nbsp;,&nbsp;,&nbsp;,NOMINA DEL," & DisplayNumericDateFromSerialNumber(asPayrollIDs(0)) & ",NOMINA DEL," & DisplayNumericDateFromSerialNumber(asPayrollIDs(1)) & ",DIFERENCIA,&nbsp;,&nbsp;", ",")
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles2, asCellWidths, asTableColors, sErrorDescription)
				asColumnsTitles = Split("TIPO,COPTO,NOMBRE DEL CONCEPTO,IMPORTE,REGS,IMPORTE,REGS,IMPORTE,REGS,PORC", ",")
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split("CENTER,LEFT,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					iConceptID = CInt(oRecordset.Fields("ConceptID").Value)
					If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
						sRowContents = "P"
					Else
						sRowContents = "D"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents &CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(0) & " Where (Payroll_" & asPayrollIDs(0) & ".ConceptID =" & iConceptID & ")", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
					If lErrorNumber = 0 Then
						If Not oRecordset1.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset1.Fields("TotalAmount").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset1.Fields("TotalPayments").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							dTotalAmount1 = CDbl(oRecordset1.Fields("TotalAmount").Value)
							dTotalPayments1 = CDbl(oRecordset1.Fields("TotalPayments").Value)
						Else
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							dTotalAmount1 = 0
							dTotalPayments1 = 0
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(1) & " Where (Payroll_" & asPayrollIDs(1) & ".ConceptID =" & iConceptID & ")", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset2)
					If lErrorNumber = 0 Then
						If Not oRecordset2.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset2.Fields("TotalAmount").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset2.Fields("TotalPayments").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							dTotalAmount2 = CDbl(oRecordset2.Fields("TotalAmount").Value)
							dTotalPayments2 = CDbl(oRecordset2.Fields("TotalPayments").Value)
						Else
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							dTotalAmount2 = 0
							dTotalPayments2 = 0
						End If
					End If
					If lErrorNumber = 0 Then
						sRowContents = sRowContents & FormatNumber((dTotalAmount1-dTotalAmount2), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((dTotalPayments1-dTotalPayments2), 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						If dTotalAmount1 <> 0 Then
							sRowContents = sRowContents & FormatNumber(((dTotalAmount1-dTotalAmount2)/dTotalAmount1)*100, 2, True, False, True)
						Else
							If dTotalAmount2 <> 0 Then
								sRowContents = sRowContents & FormatNumber(((dTotalAmount1-dTotalAmount2)/dTotalAmount2)*100, 2, True, False, True)
							Else
								sRowContents = sRowContents & "0.00"
							End If
						End If	
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					oRecordset1.Close
					oRecordset2.Close
					'If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oRecordset.Close	
			End If	
		End If	
	End If

	Set oRecordset = Nothing
	BuildReport1012 = lErrorNumber
	Err.Clear		
End Function

Function BuildReport1013(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Diferencias de empleados por unidad administrativa. Reporte basado en la hoja 001172
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1013"
	Dim sHeaderContents
	Dim iMonth
	Dim iMonthPrev
	Dim iYear
	Dim iYearPrev
	Dim iIndex
	Dim asPayrollIDs
	Dim lPayments
	Dim dPayments
	Dim sQuery
	Dim oRecordset
	Dim oAreasRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iAreaID
	Dim lTotalEmployees1
	Dim lTotalEmployees2
	Dim iCount
	Dim sTitles

	iMonth = CInt(Mid(oRequest.Item("PayrollID").Item, 5, 2))
	iYear = CInt(Mid(oRequest.Item("PayrollID").Item, 1, 4))
	If iMonth = 12 Then
		iMonth = 1
		iYear = iYear + 1
	Else
		iMonth = iMonth + 1
	End If

	sQuery = "Select PayrollID From Payrolls Where (PayrollDate>=" & oRequest.Item("PayrollID").Item & ") And (PayrollDate<=" & (iYear & Right(("0" & iMonth), Len("00"))) & "99) And (PayrollTypeID = 1) Order By PayrollID"

	sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	asPayrollIDs = ""
	iCount = 0
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			iCount = iCount + 1
			asPayrollIDs = asPayrollIDs & CStr(oRecordset.Fields("PayrollID").Value) & ";"
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	oRecordset.Close
	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1013.htm"), sErrorDescription)
	If (Len(asPayrollIDs) > 0) Then
		asPayrollIDs = Left(asPayrollIDs, (Len(asPayrollIDs) - Len(";")))
		asPayrollIDs = Split(asPayrollIDs, ";")
	End If	
	
	If (Len(sHeaderContents) > 0) Then
		sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
		sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE1 />", DisplayNumericDateFromSerialNumber(asPayrollIDs(0)))
		sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE2 />", DisplayNumericDateFromSerialNumber(asPayrollIDs(1)))
		Response.Write sHeaderContents
		
		sErrorDescription = "No se pudieron obtener los conceptos de pago."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaCode, AreaName From Areas Where (ParentID = -1) And (AreaID>0)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAreasRecordset)
		If lErrorNumber = 0 Then
			If Not oAreasRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				sTitles = "AREA,NOMBRE DE LA DELEGACIÓN,REGISTROS<BR />" & DisplayNumericDateFromSerialNumber(asPayrollIDs(0)) & ",REGISTROS<BR />" & DisplayNumericDateFromSerialNumber(asPayrollIDs(1)) & ",DIFER"
				asColumnsTitles = Split(sTitles, ",")
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split("CENTER,LEFT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oAreasRecordset.EOF
					iAreaID = CInt(oAreasRecordset.Fields("AreaID").Value)
					sRowContents = CleanStringForHTML(CStr(oAreasRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents &CleanStringForHTML(CStr(oAreasRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select count(*) As TotalEmployees From Employees, Areas, Payroll_" & asPayrollIDs(0) & ", Jobs Where (Payroll_" & asPayrollIDs(0) & ".ConceptID = -1) And (Payroll_" & asPayrollIDs(0) & ".EmployeeID = Employees.EmployeeID) And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.AreaPath Like '" & S_WILD_CHAR & "," & CLng(oAreasRecordset.Fields("AreaID").Value) & "," & S_WILD_CHAR & "')", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalEmployees").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							lTotalEmployees1 = CLng(oRecordset.Fields("TotalEmployees").Value)
						Else
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							lTotalEmployees1 = 0
						End If
					End If
					oRecordset.Close
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select count(*) As TotalEmployees From Employees, Areas, Payroll_" & asPayrollIDs(1) & ", Jobs Where (Payroll_" & asPayrollIDs(1) & ".ConceptID = -1) And (Payroll_" & asPayrollIDs(1) & ".EmployeeID = Employees.EmployeeID) And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.AreaPath Like '" & S_WILD_CHAR & "," & CLng(oAreasRecordset.Fields("AreaID").Value) & "," & S_WILD_CHAR & "')", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalEmployees").Value), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR
							lTotalEmployees2 = CLng(oRecordset.Fields("TotalEmployees").Value)
						Else
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & TABLE_SEPARATOR
							lTotalEmployees2 = 0
						End If
					End If
					oRecordset.Close
					
					If lErrorNumber = 0 Then
						sRowContents = sRowContents & FormatNumber((lTotalEmployees1-lTotalEmployees2), 0, True, False, True)
					End If
					
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oAreasRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oAreasRecordset.Close
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oAreasRecordset.Close	
			End If	
		End If	
	End If

	Set oAreasRecordset = Nothing
	BuildReport1013 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1014(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditoria de nómina. Comparativo de nóminas. Resumen de altas por unidad administrativa. Reporte basado en la hoja 001173.
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1014"
	Dim sHeaderContents
	Dim asAreas
	Dim asPayrolls
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim iIndex

	lTotal = 0

	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaShortName, AreaName From Areas Where (ParentID=-1) And AreaID>0", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows
	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1014.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"

				asColumnsTitles = Split("AREA,NOMBRE DE LA DELEGACIÓN,REGISTROS", ",")
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split("CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(CStr(asAreas(1,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(asAreas(2,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total From EmployeesHistoryListForPayroll EHLFP, Employees, Positions, Areas Where (EHLFP.EmployeeID = Employees.EmployeeID) And (EHLFP.PositionID = Positions.PositionID) And (EHLFP.AreaID = Areas.AreaID) And (Areas.AreaPath Like ',-1," & CLng(asAreas(0,iIndex)) & "," & S_WILD_CHAR & "') And (EHLFP.PayrollID = " & asPayrolls(0,0) & ") And (EHLFP.EmployeeID Not In (Select Distinct EmployeeID From EmployeesHistoryListForPayroll Where PayrollID = " & asPayrolls(0,1) & ")) ", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Total").Value), 0, True, False, True)
							lTotal = lTotal + CDbl(oRecordset.Fields("Total").Value)
						Else
							sRowContents = sRowContents & "0"
						End If
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.Close
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "TOTAL DE ALTAS DEL MES"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotal
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			End If
		End If
	End If
	
	Set oRecordset = Nothing
	BuildReport1014 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1015(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditoría de nómina. Comparativo de nóminas. Altas por unidad administrativa. Reporte basado en la hoja 001174
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1015"
	Dim sHeaderContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim asAreas
	Dim asPayrolls
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim lTotalRecords
	Dim iIndex

	lTotal = 0
	lTotalRecords = 0

	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaName From Areas Where (ParentID=-1) And (AreaID>0)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows
	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1015.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asCellAlignments = Split("CENTER,LEFT,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(CStr(asAreas(1,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(asAreas(2,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EHLFP.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Positions.PositionShortName From EmployeesHistoryListForPayroll EHLFP, Employees, Positions, Areas Where (EHLFP.EmployeeID = Employees.EmployeeID) And (EHLFP.PositionID = Positions.PositionID) And (EHLFP.AreaID = Areas.AreaID) And (Areas.AreaPath Like ',-1," & asAreas(0,iIndex) & ",%') And (EHLFP.PayrollID = " & asPayrolls(0,0) & ") And (EHLFP.EmployeeID Not In (Select Distinct EmployeeID From EmployeesHistoryListForPayroll Where PayrollID= " & asPayrolls(0,1) & ")) Order By EHLFP.EmployeeNumber", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						lTotal = 0
						If Not oRecordset.EOF Then
							Do While Not oRecordset.EOF
								sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								Else
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								lTotal = lTotal + 1
								lTotalRecords = lTotalRecords + 1
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & lTotal
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Else
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
					End If
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "<B>TOTAL DE ALTAS</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotalRecords
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oAreaRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1015 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1016(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditoría de nómina. Bajas por unidad administrativa. Reporte basado en la hoja 001176
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1016"
	Dim sHeaderContents
	Dim oRecordset
	Dim asAreas
	Dim asPayrolls
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim lTotalRecords
	Dim iIndex

	lTotal = 0
	lTotalRecords = 0
	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaShortName, AreaName From Areas Where (ParentID=-1) And AreaID>0", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows
	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1014.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"

				asColumnsTitles = Split("AREA,NOMBRE DE LA DELEGACIÓN,REGISTROS", ",")
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split("CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(CStr(asAreas(1,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(asAreas(2,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total From EmployeesHistoryListForPayroll EHLFP, Employees, Positions, Areas Where (EHLFP.EmployeeID = Employees.EmployeeID) And (EHLFP.PositionID = Positions.PositionID) And (EHLFP.AreaID = Areas.AreaID) And (Areas.AreaPath Like ',-1," & CLng(asAreas(0,iIndex)) & "," & S_WILD_CHAR & "') And (EHLFP.PayrollID = " & asPayrolls(0,1) & ") And (EHLFP.EmployeeID Not In (Select Distinct EmployeeID From EmployeesHistoryListForPayroll Where PayrollID = " & asPayrolls(0,0) & ")) ", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Total").Value), 0, True, False, True)
							lTotal = lTotal + CDbl(oRecordset.Fields("Total").Value)
						Else
							sRowContents = sRowContents & "0"
						End If
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.Close
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "TOTAL DE ALTAS DEL MES"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotal
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1016 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1017(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditorías de nómina. Bajas por unidad administrativa. Reporte basado en la hoja 001175
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1017"
	Dim sHeaderContents
	Dim iMonth
	Dim iYear
	Dim oRecordset
	Dim asAreas
	Dim asPayrolls
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim lTotalRecords
	Dim iIndex

	lTotal = 0
	lTotalRecords = 0

	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaName From Areas Where (ParentID=-1) And (AreaID>0)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows
	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1017.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asCellAlignments = Split("CENTER,LEFT,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(CStr(asAreas(1,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(asAreas(2,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EHLFP.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Positions.PositionShortName From EmployeesHistoryListForPayroll EHLFP, Employees, Positions, Areas Where (EHLFP.EmployeeID = Employees.EmployeeID) And (EHLFP.PositionID = Positions.PositionID) And (EHLFP.AreaID = Areas.AreaID) And (Areas.AreaPath Like ',-1," & asAreas(0,iIndex) & ",%') And (EHLFP.PayrollID = " & asPayrolls(0,1) & ") And (EHLFP.EmployeeID Not In (Select Distinct EmployeeID From EmployeesHistoryListForPayroll Where PayrollID= " & asPayrolls(0,0) & ")) Order By EHLFP.EmployeeNumber", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						lTotal = 0
						If Not oRecordset.EOF Then
							Do While Not oRecordset.EOF
								sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								Else
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								lTotal = lTotal + 1
								lTotalRecords = lTotalRecords + 1
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & lTotal
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Else
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
					End If
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "<B>TOTAL DE ALTAS</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotalRecords
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oAreaRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1017 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1018(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditoría de nóminas. Diferencias de sueldo por unidad administrativa. Reporte basado en la hoja 001177
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1018"
	Dim sHeaderContents
	Dim oRecordset
	Dim asAreas
	Dim asPayrolls
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim lTotalRecords
	Dim iIndex

	lTotal = 0
	lTotalRecords = 0	
	lTotal = 0
	lTotalRecords = 0
	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaShortName, AreaName From Areas Where (ParentID=-1) And AreaID>0", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows

	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1018.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asCellAlignments = Split("CENTER,LEFT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(CStr(asAreas(1,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(asAreas(2,iIndex)))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & "Sueldo Anterior"
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & "Sueldo Actual"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Actual.EmployeeID, EmployeeLastName, EmployeeLastName2, EmployeeName, pr_Ant.ConceptAmount Sueldo_Anterior, pr_Act.ConceptAmount Sueldo_Actual From EmployeesHistoryListForPayroll Previous, EmployeesHistoryListForPayroll Actual, Employees, Areas, Payroll_" & asPayrolls(0,1) & " pr_Ant, payroll_" & asPayrolls(0,0) & " pr_Act Where (Actual.EmployeeID = Previous.EmployeeID) And (Actual.PayrollID = '" & asPayrolls(0,0) & "') And (Previous.PayrollID = '" & asPayrolls(0,1) & "') And (Actual.AreaID = Previous.AreaID) And (Actual.AreaID = Areas.AreaID) And (Areas.AreaPath Like ',-1," & asAreas(0,iIndex) & ",%') And (Actual.PositionID = Previous.PositionID) And (Actual.EmployeeID = Previous.EmployeeID) And (Actual.EmployeeID = pr_Ant.EmployeeID) And (pr_Ant.EmployeeID = pr_Act.EmployeeID) And (Actual.EmployeeID = Employees.EmployeeID) And (pr_Act.ConceptID = pr_Ant.ConceptID) And (pr_Act.ConceptID = 1) And (pr_Act.ConceptAmount <> pr_Ant.ConceptAmount) Order By Actual.EmployeeID Asc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						lTotal = 0
						If Not oRecordset.EOF Then
							Do While Not oRecordset.EOF
								sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								Else
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Anterior").Value), 2, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Actual").Value), 2, True, False, True)
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								lTotal = lTotal + 1
								lTotalRecords = lTotalRecords + 1
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & lTotal
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Else
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
					End If
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "<B>TOTAL DE ALTAS</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotalRecords
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oAreaRecordset.Close
			End If
		End If
	End If	
	Set oRecordset = Nothing
	BuildReport1018 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1019(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditoría de nóminas. Diferencias de cambios de puesto por unidad administrativa. Reporte basado en la hoja 001178
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1019"
	Dim sHeaderContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lDateFrom
	Dim lDateTo
	Dim lTotal
	Dim lTotalRecords
	Dim asAreas
	Dim asPayrolls
	Dim iIndex

	lTotal = 0
	lTotalRecords = 0
	sErrorDescription = "No se pudieron obtener las delegaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaShortName, AreaName From Areas Where (ParentID=-1) And AreaID>0", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asAreas = oRecordset.GetRows
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 2 PayrollID From Payrolls Where (PayrollID <= " & oRequest.Item("PayrollID").Item & ") And (PayrollTypeID = 1) Order By 1 Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows	
	If lErrorNumber = 0 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1019.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
			sHeaderContents = Replace(sHeaderContents, "<MONTH />", UCase(asMonthNames_es(iMonth)))
			sHeaderContents = Replace(sHeaderContents, "<YEAR />", iYear)
			Response.Write sHeaderContents
			If Not oAreaRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asCellAlignments = Split("CENTER,LEFT,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asAreas,2)
					sRowContents = CleanStringForHTML(asAreas(1,iIndex))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(asAreas(2,iIndex))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & "Puesto Anterior"
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & "Puesto Actual"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asRowContents, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asRowContents, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Actual.EmployeeID, EmployeeLastName, EmployeeLastName2, EmployeeName, Pto_Anterior.PositionShortName Puesto_Anterior, Pto_Nuevo.PositionShortName Puesto_Actual From EmployeesHistoryListForPayroll Previous, EmployeesHistoryListForPayroll Actual, Positions Pto_Nuevo, Positions Pto_Anterior, Employees, Areas Where (Actual.EmployeeID = Previous.EmployeeID) And (Actual.PayrollID = '" & asPayrolls(0,0) & "') And (Previous.PayrollID = '" & asPayrolls(0,1) & "') And (Areas.AreaPath Like ',-1," & CLng(asAreas(0,iIndex)) & "," & S_WILD_CHAR & "') And (Actual.PositionID <> Previous.PositionID) And (Actual.PositionID = Pto_Nuevo.PositionID) And (Actual.EmployeeID = Employees.EmployeeID) And (Previous.PositionID = Pto_Anterior.PositionID) And (Actual.AreaID = Areas.AreaID) Order By Actual.EmployeeID Asc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						lTotal = 0
						If Not oRecordset.EOF Then
							Do While Not oRecordset.EOF
								sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								Else
									sRowContents = sRowContents & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Puesto_Anterior").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Puesto_Actual").Value))
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								lTotal = lTotal + 1
								lTotalRecords = lTotalRecords + 1
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & lTotal
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Else
							sRowContents = TABLE_SEPARATOR
							sRowContents = sRowContents & "<B>Registros</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "0"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
					End If
				Next
				sRowContents = TABLE_SEPARATOR
				sRowContents = sRowContents & "<B>TOTAL DE CAMBIOS</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR
				sRowContents = sRowContents & lTotalRecords
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oAreaRecordset.Close
			End If
		End If
	End If
	
	Set oRecordset = Nothing
	BuildReport1019 = lErrorNumber
	Err.Clear
End Function 

%>