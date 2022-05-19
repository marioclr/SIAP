<%
Function BuildReports1003Cancel(oRequest, oADODBConnection, bReview, sErrorDescription)
'************************************************************
'Purpose: Listado de firmas masivo. Reporte basado en la hoja 001157
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bReview
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReports1003Cancel"
	Const N_EMPLOYEES_PER_FILE = 2000
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
	Dim lCurrentForPayrollDate
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
	Else
        If StrComp(oRequest("EmployeeTypeID").Item, "7", vbBinaryCompare) = 0 Then
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
			If oRequest("EmployeeTypeID").Item <> 7 Then
				lStartPayrollDate = GetPayrollStartDate(lForPayrollID)
			Else
				lStartPayrollDate = CLng(Left(lForPayrollID, Len("000000")) & "01")
			End If
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
				sOrderBy = "Payrolls.ForPayrollDate, CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, OrderInList, RecordDate, RecordID"
			End If
			sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
			If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
				sOrderBy = Replace(sOrderBy, "EmployeesHistoryList.EmployeeNumber", "EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber")
				If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
					sOrderBy = "CheckNumber, " & sOrderBy
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, Payments, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, Payments, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesBeneficiariesLKP.PaymentCenterID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, BeneficiaryName, BeneficiaryLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End BeneficiaryLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesBeneficiariesLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
				sOrderBy = Replace(sOrderBy, "EmployeesHistoryList.EmployeeNumber", "EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber")
				If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
					sOrderBy = "CheckNumber, " & sOrderBy
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesHistoryListForPayroll, EmployeesCreditorsLKP, Employees, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Payroll_" & lPayrollID & ", Concepts, Payments, EmployeesChangesLKP As EmpChLKP, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payments.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lPayrollID & ".RecordID) And (EmployeesCreditorsLKP.EmployeeID = EmpChLKP.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Payroll_" & lPayrollID & ".ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (EmployeesHistoryListForPayroll.PayrollID = " & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.AreaID = Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (PaymentCenters.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate <= " & lPayrollID & ") And (EmployeesCreditorsLKP.EndDate >= " & lPayrollID & ") And (companies.StartDate <= " & lPayrollID & ") And (Companies.EndDate >= " & lPayrollID & ") And (Zones.StartDate<=" & lPayrollID & ") And (Zones.EndDate>=" & lPayrollID & ") And (PaymentCenters.StartDate <= " & lPayrollID & ") And (PaymentCenters.EndDate >= " & lPayrollID & ") And (Areas.StartDate <= " & lPayrollID & ") And (Areas.EndDate >= " & lPayrollID & ") And (Positions.StartDate <= " & lPayrollID & ") And (Positions.EndDate >= " & lPayrollID & ") And (Levels.StartDate <= " & lPayrollID & ") And (Levels.EndDate >= " & lPayrollID & ") And (GroupGradeLevels.StartDate <= " & lPayrollID & ") And (GroupGradeLevels.EndDate >= " & lPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesHistoryListForPayroll, EmployeesCreditorsLKP, Employees, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Payroll_" & lPayrollID & ", Concepts, Payments, EmployeesChangesLKP As EmpChLKP, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payments.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lPayrollID & ".RecordID) And (EmployeesCreditorsLKP.EmployeeID = EmpChLKP.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Payroll_" & lPayrollID & ".ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (EmployeesHistoryListForPayroll.PayrollID = " & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.AreaID = Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (PaymentCenters.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate <= " & lPayrollID & ") And (EmployeesCreditorsLKP.EndDate >= " & lPayrollID & ") And (companies.StartDate <= " & lPayrollID & ") And (Companies.EndDate >= " & lPayrollID & ") And (Zones.StartDate<=" & lPayrollID & ") And (Zones.EndDate>=" & lPayrollID & ") And (PaymentCenters.StartDate <= " & lPayrollID & ") And (PaymentCenters.EndDate >= " & lPayrollID & ") And (Areas.StartDate <= " & lPayrollID & ") And (Areas.EndDate >= " & lPayrollID & ") And (Positions.StartDate <= " & lPayrollID & ") And (Positions.EndDate >= " & lPayrollID & ") And (Levels.StartDate <= " & lPayrollID & ") And (Levels.EndDate >= " & lPayrollID & ") And (GroupGradeLevels.StartDate <= " & lPayrollID & ") And (GroupGradeLevels.EndDate >= " & lPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Payroll_" & lPayrollID & ".ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.concepts40 From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP EmpChLKP, Companies, Zones, Areas As PaymentCenters, ZoneTypes, Areas, Positions, Levels, GroupGradeLevels, Zones As AreasZones, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID = EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesCreditorsLKP.CreditorNumber = Payroll_" & lForPayrollID & ".EmployeeID) And (EmployeesCreditorsLKP.EmployeeID = Payroll_" & lForPayrollID & ".RecordID)  And (EmployeesHistoryListForPayroll.EmployeeID = EmpChLKP.EmployeeID) And (employeesHistoryListForPayroll.CompanyID = Companies.CompanyID) And (EmployeesHistoryListForPayroll.ZoneID = Zones.ZoneID) And (Zones.ParentID = AreasZones.ZoneID) And (AreasZones.ParentID = ParentZones.ZoneID) And (Zones.ZoneTypeID = ZoneTypes.ZoneTypeID) And (EmployeesHistoryListForPayroll.AreaID = PaymentCenters.AreaID) And (PaymentCenters.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID = Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID = Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (Payroll_20120630.ConceptID = Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PayrollID = " & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate <= " & lForPayrollID & ")  And (EmployeesCreditorsLKP.EndDate >= " & lForPayrollID & ") And (companies.StartDate <= " & lForPayrollID & ") And (Companies.EndDate >= " & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate <= " & lForPayrollID & ") And (PaymentCenters.EndDate >= " & lForPayrollID & ") And (Areas.StartDate <= " & lForPayrollID & ") And (Areas.EndDate >= " & lForPayrollID & ") And (Positions.StartDate <= " & lForPayrollID & ") And (Positions.EndDate >= " & lForPayrollID & ") And (Levels.StartDate <= " & lForPayrollID & ") And (Levels.EndDate >= " & lForPayrollID & ") And (GroupGradeLevels.StartDate <= " & lForPayrollID & ") And (GroupGradeLevels.EndDate >= " & lForPayrollID & ") And (EmpChLKP.PayrollID = " & lForPayrollID & ") And (EmpChLKP.PayrollDate = Payroll_" & lForPayrollID & ".RecordDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"ParentAreas.AreaCode,",""), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, '0' As EmployeeTypeID, EmployeesCreditorsLKP.PaymentCenterID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, CreditorName, CreditorLastName, Case When CreditorLastName2 Is Null Then ' ' Else CreditorLastName2 End CreditorLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, '----------' As CheckNumber, BankAccounts.AccountNumber, BankAccounts.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From EmployeesCreditorsLKP, BankAccounts, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesCreditorsLKP.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Employees.EmployeeID) And (EmpChLKP.PayrollID=" & lPayrollID & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				End If
			Else
				If InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Payrolls, Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payments, Payroll_" & lPayrollID & ", Payrolls, Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Payrolls, Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas.StartDate<=Payrolls.ForPayrollDate) And (Areas.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (GroupGradeLevels.StartDate<=Payrolls.ForPayrollDate) And (GroupGradeLevels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Payrolls.ForPayrollDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName, Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, Areas.EconomicZoneID, EmployeesHistoryListForPayroll.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, ConceptAmount, '----------' As CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.BankID, EmpChLKP.FirstDate, EmpChLKP.LastDate, EmpChLKP.Concepts40 From Payroll_" & lPayrollID & ", Payrolls, Concepts, Employees, EmployeesHistoryListForPayroll, EmployeesChangesLKP As EmpChLKP, Companies, Areas, Areas As ParentAreas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmpChLKP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmpChLKP.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmpChLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (PaymentCenters.CompanyID=Companies.CompanyID) And (PaymentCenters.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (PaymentCenters.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas.StartDate<=Payrolls.ForPayrollDate) And (Areas.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (GroupGradeLevels.StartDate<=Payrolls.ForPayrollDate) And (GroupGradeLevels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By " & Replace(Replace(sOrderBy, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " -->" & vbNewLine
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
                            lCurrentForPayrollDate = -1
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
									If (Len(sContents) > 0) And ((lCurrentForPayrollDate <> CLng(oRecordset.Fields("ForPayrollDate").Value)) Or (lCurrentCompanyID <> CLng(oRecordset.Fields("CompanyID").Value)) Or (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Or (StrComp(sCurrentEmployeeID, sTempCurrent, vbBinaryCompare) <> 0)) Then
										If lTempID > -1 Then
											If CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 7 Then
												If lTempDeduction = 0 Then
													asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
												Else
													asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
												End If
											Else
												If lTempDeduction = 0 Then
													asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
													dTempAmount = 0
												Else
													asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & sTempShortName & SECOND_LIST_SEPARATOR & dTempAmount & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
													dTempAmount = 0
												End If
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
										If CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 7 Then
											iBound = Int(UBound(asConceptsP) / 2)
											jBound = Int(UBound(asConceptsD) / 2)
										Else
											iBound = Int(UBound(asConceptsP))
											jBound = Int(UBound(asConceptsD))
										End If
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
									If (lCurrentForPayrollDate <> CLng(oRecordset.Fields("ForPayrollDate").Value)) Or (lCurrentCompanyID <> CLng(oRecordset.Fields("CompanyID").Value)) Or (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
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
                                        'lCurrentForPayrollDate = CLng(oRecordset.Fields("ForPayrollDate").Value)
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
									If (lCurrentForPayrollDate <> CLng(oRecordset.Fields("ForPayrollDate").Value)) Or (StrComp(sCurrentEmployeeID, sTempCurrent, vbBinaryCompare) <> 0) Then
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
                                            'lCurrentForPayrollDate = CLng(oRecordset.Fields("ForPayrollDate").Value)
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
                                        lCurrentForPayrollDate = CLng(oRecordset.Fields("ForPayrollDate").Value)
										sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
										If bReview Then sCurrentEmployeeID = sCurrentEmployeeID & "," & CStr(oRecordset.Fields("RecordDate").Value)
									End If
									Select Case CLng(oRecordset.Fields("ConceptID").Value)
										Case -2
											adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(2) = adURTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(2) = adCTTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case -1
											adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(1) = adURTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(1) = adCTTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case 0
											adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adURTotal(0) = adURTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
											adCTTotal(0) = adCTTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
										Case Else
											If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
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
													If CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 7 Then
														lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)
													Else
														'lTempStartDate = lStartPayrollDate
                                                            lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)  
													End If
													If lTempStartDate = 0 Then lTempStartDate = lPayrollID
												End If
												dTempAmount = dTempAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
												lTempEndDate = CLng(oRecordset.Fields("LastDate").Value)
												If lTempEndDate = 0 Then lTempEndDate = lPayrollID

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
													If CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 7 Then
														lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)
														If lTempStartDate = 0 Then lTempStartDate = lPayrollID
													Else
														'lTempStartDate = lStartPayrollDate
                                                        lTempStartDate = CLng(oRecordset.Fields("FirstDate").Value)
													End If
												End If
												dTempAmount = dTempAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
												lTempEndDate = CLng(oRecordset.Fields("LastDate").Value)
												If lTempEndDate = 0 Then lTempEndDate = lPayrollID
											End If
									End Select

									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								'oRecordset.Close
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
								If CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 7 Then
									iBound = Int(UBound(asConceptsP) / 2)
									jBound = Int(UBound(asConceptsD) / 2)
								Else
									iBound = Int(UBound(asConceptsP))
									jBound = Int(UBound(asConceptsD))
								End If
								oRecordset.Close
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
	BuildReports1003Cancel = lErrorNumber
	Err.Clear
End Function

Function BuildReport1006Cancel(oRequest, oADODBConnection, bReview, sErrorDescription)
'************************************************************
'Purpose: Listado de firmas masivo. Reporte basado en la hoja 001157
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bReview
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1006Cancel"
	Dim sDistinct
    Dim sQueryBegin
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

	sQueryBegin = ""
    Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
    If InStr(1, sCondition, " And (EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
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
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
	ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
	Else
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.IsDeduction=0) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.IsDeduction=0) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
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
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
				ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.IsDeduction=1) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.IsDeduction=1) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					'Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.IsDeduction=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
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
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
				ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, BankAccounts, EmployeesHistoryListForPayroll, Companies, Areas, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.ConceptID=0) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Payrolls, Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=Payrolls.ForPayrollDate) And (Concepts.EndDate>=Payrolls.ForPayrollDate) And (Concepts.ConceptID=0) And (Companies.StartDate<=Payrolls.ForPayrollDate) And (Companies.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (Zones.StartDate<=Payrolls.ForPayrollDate) And (Zones.EndDate>=Payrolls.ForPayrollDate) And (Positions.StartDate<=Payrolls.ForPayrollDate) And (Positions.EndDate>=Payrolls.ForPayrollDate) And (EmployeeTypes.StartDate<=Payrolls.ForPayrollDate) And (EmployeeTypes.EndDate>=Payrolls.ForPayrollDate) And (PositionTypes.StartDate<=Payrolls.ForPayrollDate) And (PositionTypes.EndDate>=Payrolls.ForPayrollDate) And (Levels.StartDate<=Payrolls.ForPayrollDate) And (Levels.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					'Response.Write vbNewLine & "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName, Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter, Positions, EmployeeTypes, PositionTypes, Levels, Areas As PaymentCenters, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID=0) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Concepts.ConceptID, ConceptShortName, ConceptName, Budgets1.BudgetShortName Order by Concepts.ConceptID -->" & vbNewLine
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
	BuildReport1006Cancel = lErrorNumber
	Err.Clear
End Function

Function BuildReport1490Cancel(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll group by states and filtered
'         by banks
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1490Cancel"
	Dim sContents
	Dim sCondition
	Dim sDistinct
	Dim sField
	Dim sField2
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim oRecordset
	Dim adTotals
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

	If Len(oRequest("ZoneID").Item) > 0 Then
		asStateIDs = Replace(oRequest("ZoneID").Item, " ", "")
		asStateIDs = asStateIDs & ",1000"
		If (InStr(1, asStateIDs, ",9,", vbBinaryCompare) > 0) Then
			asStateIDs = Replace(asStateIDs, ",9,", ",") & ",9"
		ElseIf (InStr(1, asStateIDs, "9,", vbBinaryCompare) = 1) Then
			asStateIDs = Replace(asStateIDs, "9,", "", 1, 1, vbBinaryCompare) & ",9"
		ElseIf (InStr(1, asStateIDs, ",9", vbBinaryCompare) = (Len(asStateIDs) - 1)) Then
		Else
			asStateIDs = asStateIDs & ",1000"
		End If
		asStateIDs = Split(asStateIDs, ",")
	Else
		asStateIDs = Split(S_STATE_IDS, ",")
	End If
	adTotals = Split(",,,", ",")
	For iIndex = 0 To UBound(adTotals)
		adTotals(iIndex) = Split(",", ",")
		adTotals(iIndex)(0) = 0
		adTotals(iIndex)(1) = 0
	Next
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

	If (iConnectionType <> ACCESS) And (iConnectionType <> ACCESS_DSN) Then
		sDistinct = "Distinct "
	Else
		sDistinct = ""
	End If
	If Len(oRequest("ForWorkingCenter").Item) = 0 Then
		sField = "ZonesForPaymentCenter"
		sField2 = "PaymentCenters"
	Else
		sField = "Zones"
		sField2 = "Areas2"
	End If
	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1490.htm"), sErrorDescription)
	sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
	If Len(oRequest("ForWorkingCenter").Item) = 0 Then
		sContents = Replace(sContents, "<WORKING_CENTER />", "POR CENTRO DE PAGO")
	Else
		sContents = Replace(sContents, "<WORKING_CENTER />", "PRESUPUESTAL")
	End If
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
		asColumnsTitles = Split("Foráneo,Registros,Percepciones,Deducciones,Líquido", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "0", vbBinaryCompare) = 0) Then
			For iIndex = 0 To UBound(asStateIDs) - 1
				If CLng(asStateIDs(iIndex)) = -1 Then
					sStateName = "20A HOSP. REG. PDTE. JUAREZ OAXACA, OAX."
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
					Else
					    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID=38) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					    Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID=38) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
					End If
				Else
					Call GetNameFromTable(oADODBConnection, "States", asStateIDs(iIndex), "", "", sStateName, "")
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
					End If
				End If
				If lErrorNumber = 0 Then
					For jIndex = 0 To UBound(adTotals)
						adTotals(jIndex)(0) = 0
					Next
					If Not oRecordset.EOF Then
						adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
						Do While Not oRecordset.EOF
							If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
								adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
								adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							Else
								adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
							End If
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
					End If
					sRowContents = CleanStringForHTML(sStateName)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
					adTotals(3)(1) = adTotals(3)(1) + adTotals(3)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
					adTotals(1)(1) = adTotals(1)(1) + adTotals(1)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
					adTotals(0)(1) = adTotals(0)(1) + adTotals(0)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
					adTotals(2)(1) = adTotals(2)(1) + adTotals(2)(0)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
			Next

			sRowContents = "<B>TOTAL FORÁNEOS</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(1), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(1), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(1), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(1), 2, True, False, True) & "</B>"
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

		For iIndex = 0 To UBound(adTotals)
			adTotals(iIndex)(0) = 0
		Next
		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "1", vbBinaryCompare) = 0) Then
		If StrComp(asStateIDs(UBound(asStateIDs)), "9", vbBinaryCompare) = 0 Then
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
			End If
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						Do While Not oRecordset.EOF
							adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
					End If
					oRecordset.Close
				End If
			End If
			sRowContents = "LOCAL" & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "<B>TOTAL LOCAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(0), 2, True, False, True) & "</B>"
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

			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "CHEQUE"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (EmployeesHistoryListForPayroll.AccountNumber<>'.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (EmployeesHistoryListForPayroll.AccountNumber<>'.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "DÉBITO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			For iIndex = 0 To UBound(adTotals)
				adTotals(iIndex)(0) = 0
			Next
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", Payrolls, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payrolls.PayrollID=Payroll_" & lPayrollID & ".RecordID) And And (EmployeesHistoryListForPayroll.PayrollID=Payroll_" & lPayrollID & ".RecordID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=Payrolls.ForPayrollDate) And (Areas1.EndDate>=Payrolls.ForPayrollDate) And (Areas2.StartDate<=Payrolls.ForPayrollDate) And (Areas2.EndDate>=Payrolls.ForPayrollDate) And (PaymentCenters.StartDate<=Payrolls.ForPayrollDate) And (PaymentCenters.EndDate>=Payrolls.ForPayrollDate) And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "<B>TOTAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If
		End If

		If (Len(oRequest("StateType").Item) = 0) Then
			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "FORÁNEO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(1), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(1), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(1), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(1), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "LOCAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "<B>TOTAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(1) + adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(1) + adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(1) + adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(1) + adTotals(2)(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If
	Response.Write "</TABLE>"

	Set oRecordset = Nothing
	BuildReport1490Cancel = lErrorNumber
	Err.Clear
End Function
%>