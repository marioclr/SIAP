<%
Function DisplayEmployeeConceptsTable(oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the payment concepts for the given employee
'Inputs:  oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeConceptsTable"
	Dim oRecordset
	Dim oRecordsetPay
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim dTotal
	Dim sFontBegin
	Dim sFontEnd
	Dim lErrorNumber

	If Len(oRequest("PayrollID").Item) = 0 Then
		Call GetNameFromTable(oADODBConnection, "LastPayrollID", "-1", "", "", aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), "")
	Else
		aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = CDbl(oRequest("PayrollID").Item)
	End If
	lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Concepto,Fecha de Imputación,Importe", ",", -1, vbBinaryCompare)
			asCellWidths = Split("300,200,300", ",", -1, vbBinaryCompare)
			asCellAlignments = Split(",,RIGHT", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, RecordDate,ConceptShortName, ConceptName, IsDeduction, ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ", Concepts Where (Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ".ConceptID=Concepts.ConceptID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (Concepts.StartDate<=" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ") And (Concepts.EndDate>=" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ") Order By IsDeduction, OrderInList, ConceptShortName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			Response.Write "<!-- Query: Select Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ", Concepts Where (Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ".ConceptID=Concepts.ConceptID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (Concepts.StartDate<=" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ") And (Concepts.EndDate>=" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ") Order By IsDeduction, OrderInList, ConceptShortName -->"
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					dTotal = 0
					Do While Not oRecordset.EOF
						sFontBegin = ""
						sFontEnd = ""
						If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
							dTotal = dTotal - CDbl(oRecordset.Fields("ConceptAmount").Value)
							sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
							sFontEnd = "</FONT>"
						Else
							dTotal = dTotal + CDbl(oRecordset.Fields("ConceptAmount").Value)
						End If
						If CLng(oRecordset.Fields("ConceptID").Value) <= 0 Then
							sFontBegin = sFontBegin & "<B>"
							sFontEnd = "</B>" & sFontEnd
						End If
						sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RecordDate").Value)) & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sFontEnd
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
						If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
				Else
					sRowContents = "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>0.00</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
				oRecordset.Close
'				sRowContents = "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
'				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
'				If bForExport Then
'					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
'				Else
'					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
'				End If
			End If
		Response.Write "</TABLE></DIV>"
	End If
	aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = "-2,-1,0,55," & MAIN_CONCEPTS_FOR_PAYROLL

	Set oPayrollRecordset = Nothing
	Set oRecordset = Nothing
	DisplayEmployeeConceptsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeConceptsTableSp(oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the payment concepts for the given employee
'Inputs:  oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeConceptsTableSp"
	Dim asQueries
	Dim sPeriods
	Dim iIndex
	Dim oRecordset
	Dim oPayrollRecordset
	Dim sNames
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
	Dim sAction
	Dim dTotal
	Dim dTempTotal
	Dim dSchoolarship
	Dim lStartDate
	Dim lEndDate
	Dim dHours
	Dim bDisplayForm
	Dim bSchoolarship
	Dim sConceptsIDs
	Dim sConceptsAmounts
	Dim sConceptsToDisable
	Dim bSkip
	Dim lErrorNumber

Call DisplayTimeStamp("INICIO")

	bDisplayForm = False
	If Len(oRequest("PayrollID").Item) = 0 Then
		Call GetNameFromTable(oADODBConnection, "LastPayrollID", "-1", "", "", aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), "")
	Else
		aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = CDbl(oRequest("PayrollID").Item)
	End If
	lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
				asColumnsTitles = Split("Acciones,Concepto,Importe,Mínimo,Máximo,Ausencias", ",", -1, vbBinaryCompare)
				asCellWidths = Split("140,200,150,150,150,150", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("&nbsp;,Concepto,Importe,Mínimo,Máximo,Ausencias", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,220,175,175,175,175", ",", -1, vbBinaryCompare)
			End If
			asCellAlignments = Split("CENTER,,RIGHT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			sAction = "ShowInfo"
			If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sAction = "Change"
			dTotal = 0
			dSchoolarship = 0
			bSchoolarship = False
			sConceptsIDs = ""
			sConceptsAmounts = ""
			aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = "-2,-1,0,55"
			sPeriods = ""
			sPeriods = GetPeriodsForPayroll(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), -1)
			If Len(sPeriods) > 0 Then sPeriods = " And (PeriodID In (" & sPeriods & ")) "

			If B_ISSSTE Then
				Select Case aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)
					Case 0
						asQueries = "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ") And (ConceptsValues.LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.EconomicZoneID=" & aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) & ") And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName" & LIST_SEPARATOR
					Case 1
						asQueries = "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ") And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName" & LIST_SEPARATOR
					Case 2, 4, 5
						asQueries = "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ") And (ConceptsValues.LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.EconomicZoneID=" & aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) & ") And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName" & LIST_SEPARATOR & _
									"Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.WorkingHours Not In (6.5,8)) And (ConceptsValues.LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.EconomicZoneID=" & aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) & ") And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName" & LIST_SEPARATOR
					Case 3, 6
						asQueries = "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName" & LIST_SEPARATOR
					Case Else
						asQueries = ""
				End Select
				asQueries = asQueries & "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID In (-1," & aEmployeeComponent (N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")) And (ConceptsValues.PositionTypeID In (-1," & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ")) And (ConceptsValues.JobStatusID In (-1," & aEmployeeComponent(N_JOB_STATUS_ID_EMPLOYEE) & ")) And (ConceptsValues.ClassificationID In (-1," & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ")) And (ConceptsValues.GroupGradeLevelID In (-1," & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ")) And (ConceptsValues.IntegrationID In (-1," & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ")) And (ConceptsValues.JourneyID In (-1," & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ")) And (ConceptsValues.WorkingHours In (-1," & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ")) And (ConceptsValues.LevelID In (-1," & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ")) And (ConceptsValues.EconomicZoneID In (0," & aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) & ")) And (ConceptsValues.ServiceID In (-1," & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ")) And (ConceptsValues.GenderID In (-1," & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ")) And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName, EmployeeTypeID Desc, PositionTypeID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc"
			Else
				asQueries = "Select ConceptsValues.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, QttyValues, ConceptTypes, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.ConceptQttyID=QttyValues.QttyID) And (ConceptsValues.ConceptTypeID=ConceptTypes.ConceptTypeID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (ConceptsValues.ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ") And (ConceptsValues.GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ") And (ConceptsValues.JourneyID In (-1," & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ")) And (ConceptsValues.WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ") And (ConceptsValues.LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ") And (ConceptsValues.EconomicZoneID=" & aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) & ") And (ConceptsValues.ServiceID In (-1," & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ")) And (ConceptsValues.GenderID In (-1," & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ")) And (ConceptsValues.EndDate=30000000) And (Concepts.EndDate=30000000) <CONDITION /> Order By IsDeduction, OrderInList, ConceptShortName"
			End If
			asQueries = Split(asQueries, LIST_SEPARATOR, -1, vbBinaryCompare)

			For iIndex = 0 To UBound(asQueries)
				Response.Write "<!-- QUERY: " & Replace(asQueries(iIndex), "<CONDITION />", " And (Concepts.ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "))" & sPeriods) & " -->" & vbNewLine
				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, Replace(asQueries(iIndex), "<CONDITION />", " And (Concepts.ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "))" & sPeriods), "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						bSkip = False
						If InStr(1, ("," & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & ","), ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) = 0 Then
							If (Not bSkip) And (CLng(oRecordset.Fields("AntiquityID").Value) > -1) And (aEmployeeComponent(N_START_DATE_EMPLOYEE) > 0) Then
								lStartDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("EndYears").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								lEndDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("StartYears").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								If (aEmployeeComponent(N_START_DATE_EMPLOYEE) < lStartDate) Or (aEmployeeComponent(N_START_DATE_EMPLOYEE) >= lEndDate) Then bSkip = True
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("Antiquity2ID").Value) > -1) And (aEmployeeComponent(N_START_DATE_EMPLOYEE) > 0) Then
								lStartDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("EndYears2").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								lEndDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("StartYears2").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								If (aEmployeeComponent(N_START_DATE2_EMPLOYEE) < lStartDate) Or (aEmployeeComponent(N_START_DATE2_EMPLOYEE) >= lEndDate) Then bSkip = True
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("Antiquity3ID").Value) > -1) And (aEmployeeComponent(N_START_DATE_EMPLOYEE) > 0) Then
								lStartDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("EndYears3").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								lEndDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("StartYears3").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								If (aEmployeeComponent(N_START_DATE_EMPLOYEE) < lStartDate) Or (aEmployeeComponent(N_START_DATE_EMPLOYEE) >= lEndDate) Then bSkip = True
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("Antiquity4ID").Value) > -1) And (aEmployeeComponent(N_START_DATE_EMPLOYEE) > 0) Then
								lStartDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("EndYears4").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								lEndDate = CLng(Left(GetSerialNumberForDate(DateAdd("m", (-12 * CDbl(oRecordset.Fields("StartYears4").Value)), GetDateFromSerialNumber(aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE)))), Len("00000000")))
								If (aEmployeeComponent(N_START_DATE_EMPLOYEE) < lStartDate) Or (aEmployeeComponent(N_START_DATE_EMPLOYEE) >= lEndDate) Then bSkip = True
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("ForRisk").Value) > 0) Then
								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesRisksLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (RiskLevel=" & CStr(oRecordset.Fields("ForRisk").Value) & ")", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									bSkip = oPayrollRecordset.EOF
									oPayrollRecordset.Close
								End If
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("HasChildren").Value) > 0) Then
								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate In (0,30000000))", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									bSkip = oPayrollRecordset.EOF
									oPayrollRecordset.Close
								End If
							End If
							If (Not bSkip) And (CLng(oRecordset.Fields("HasSyndicate").Value) > 0) Then
								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesSyndicatesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (SyndicateID=" & CStr(oRecordset.Fields("HasSyndicate").Value) & ") And (EndDate=30000000)", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									bSkip = oPayrollRecordset.EOF
									oPayrollRecordset.Close
								End If
							End If
							If (Not bSkip) And (InStr(1, ("," & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL & ","), ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0) And (CLng(oRecordset.Fields("SchoolarshipID").Value) > -1) Then
								bSchoolarship = True
								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(ChildID) As ChildrenCount From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (LevelID=" & CLng(oRecordset.Fields("SchoolarshipID").Value) & ") And (ChildEndDate=0)", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									If Not oPayrollRecordset.EOF Then
										dSchoolarship = dSchoolarship + CInt(oPayrollRecordset.Fields("ChildrenCount").Value) * CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
									oPayrollRecordset.Close
								End If
								bSkip = True
							ElseIf bSchoolarship Then
								bSchoolarship = False
								aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "," & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL
								If bForExport Then
									sRowContents = "IcnAlarmXXX.txt"
								Else
									sRowContents = "<IMG SRC=""Images/IcnAlarmXXX.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""XXX"" />"
								End If
								Call GetNameFromTable(oADODBConnection, "FullConcepts", SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL, "", "<BR />", sNames, sErrorDescription)
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(sNames) & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "$" & FormatNumber(dSchoolarship, 2, True, False, True) & sFontEnd & TABLE_SEPARATOR & "<CENTER>---</CENTER>" & TABLE_SEPARATOR & "<CENTER>---</CENTER>" & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
								dTotal = dTotal + dSchoolarship

								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID In (" & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL & ")) Order By RecordDate, RecordID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									If oPayrollRecordset.EOF Then
										bDisplayForm = True
										sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmYellow.gif"), "ALT=""XXX""", "ALT=""Este concepto no se ha registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>!</B>")
									ElseIf StrComp(FormatNumber(CDbl(oPayrollRecordset.Fields("ConceptAmount").Value), 2, True, False, True), FormatNumber(dSchoolarship, 2, True, False, True), vbBinaryCompare) <> 0 Then
										bDisplayForm = True
										sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmRed.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto difiere de lo registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>X</B>")
									Else
										sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmGreen.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto coincide con lo registrado la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>Ok</B>")
									End If
									oPayrollRecordset.Close
								Else
									bDisplayForm = True
								End If

								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								sConceptsIDs = sConceptsIDs & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL & LIST_SEPARATOR
								sConceptsAmounts = sConceptsAmounts & dSchoolarship & LIST_SEPARATOR
							End If
							If Not bSkip Then
								aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "," & CStr(oRecordset.Fields("ConceptID").Value)

								If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
									sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
									sFontEnd = "</FONT>"
								End If
								If bForExport Then
									sRowContents = "IcnAlarmXXX.txt"
								Else
									sRowContents = "<IMG SRC=""Images/IcnAlarmXXX.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""XXX"" />"
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
									dTempTotal = 0
									If InStr(1, ("," & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL & ","), ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0 Then
										dTempTotal = dSchoolarship
									Else
										Select Case CInt(oRecordset.Fields("ConceptQttyID").Value)
											Case 1
												dTempTotal = CDbl(oRecordset.Fields("ConceptAmount").Value)
												sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True)
											Case 2
												lErrorNumber = GetConceptAmount(oADODBConnection, aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), aEmployeeComponent(N_ID_EMPLOYEE), aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE), aEmployeeComponent(N_GEOGRAPHICAL_ZONE_ID_EMPLOYEE), aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE), CStr(oRecordset.Fields("AppliesToID").Value), True, False, dTempTotal, sErrorDescription)
												dTempTotal = dTempTotal * CDbl(oRecordset.Fields("ConceptAmount").Value) / 100
												sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True) & "<BR />"
												Call GetNameFromTable(oADODBConnection, "FullConcepts", CStr(oRecordset.Fields("AppliesToID").Value), "<BR />", "", sNames, sErrorDescription)
												sRowContents = sRowContents & "(" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value) & " sobre " & sNames) & ")"
											Case Else
												lErrorNumber = GetConceptAmount(oADODBConnection, aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), aEmployeeComponent(N_ID_EMPLOYEE), aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE), aEmployeeComponent(N_GEOGRAPHICAL_ZONE_ID_EMPLOYEE), aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE), CStr(oRecordset.Fields("ConceptID").Value), False, True, dTempTotal, sErrorDescription)
												sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True) & "<BR />"
												sRowContents = sRowContents & "($" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value)) & ")"
										End Select
									End If
									If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
										dTotal = dTotal - dTempTotal
									Else
										dTotal = dTotal + dTempTotal
									End If
								sRowContents = sRowContents & sFontEnd & TABLE_SEPARATOR & "<CENTER>---</CENTER>" & TABLE_SEPARATOR & "<CENTER>---</CENTER>" & TABLE_SEPARATOR & "<CENTER>---</CENTER>"

								If dTempTotal > 0 Then
									sErrorDescription = "No se pudo obtener la información del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & ") Order By RecordDate, RecordID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
									If lErrorNumber = 0 Then
										If oPayrollRecordset.EOF Then
											bDisplayForm = True
											sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmYellow.gif"), "ALT=""XXX""", "ALT=""Este concepto no se ha registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>!</B>")
										ElseIf StrComp(FormatNumber(CDbl(oPayrollRecordset.Fields("ConceptAmount").Value), 2, True, False, True), FormatNumber(dTempTotal, 2, True, False, True), vbBinaryCompare) <> 0 Then
											bDisplayForm = True
											sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmRed.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto difiere de lo registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>X</B>")
										Else
											sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmGreen.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto coincide con lo registrado la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>Ok</B>")
										End If
										oPayrollRecordset.Close
									Else
										bDisplayForm = True
									End If

									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									sConceptsIDs = sConceptsIDs & CLng(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR
									sConceptsAmounts = sConceptsAmounts & dTempTotal & LIST_SEPARATOR
								End If
							End If
						End If
						oRecordset.MoveNext
						If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
				End If
			Next

			sConceptsToDisable = "-760211"
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.*, ConceptShortName, ConceptName, IsDeduction, QttyValues.QttyName, ConceptTypeName, MinQttyValues.QttyName As MinQttyName, MaxQttyValues.QttyName As MaxQttyName, AbsenceTypeName From EmployeesConceptsLKP, Concepts, QttyValues, ConceptTypes, QttyValues As MinQttyValues, QttyValues As MaxQttyValues, AbsenceTypes Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.ConceptQttyID=QttyValues.QttyID) And (EmployeesConceptsLKP.ConceptTypeID=ConceptTypes.ConceptTypeID) And (EmployeesConceptsLKP.ConceptMinQttyID=MinQttyValues.QttyID) And (EmployeesConceptsLKP.ConceptMaxQttyID=MaxQttyValues.QttyID) And (EmployeesConceptsLKP.AbsenceTypeID=AbsenceTypes.AbsenceTypeID) And (EmployeesConceptsLKP.ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & ")) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeesConceptsLKP.EndDate=30000000) " & sPeriods & " Order By IsDeduction, OrderInList, ConceptShortName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "," & CStr(oRecordset.Fields("ConceptID").Value)
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ConceptID").Value), oRequest("ConceptID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
						sConceptsToDisable = sConceptsToDisable & "," & CStr(oRecordset.Fields("ConceptID").Value)
					ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					If CInt(oRecordset.Fields("Active").Value) = 1 Then
						If bForExport Then
							sRowContents = "IcnAlarmXXX.txt"
						Else
							sRowContents = "<IMG SRC=""Images/IcnAlarmXXX.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""XXX"" /><BR />"
						End If
					End If
					If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & "&nbsp;"
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeConcepts&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&Tab=3&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If B_DELETE And (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeConcepts&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&Tab=3&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								If CInt(oRecordset.Fields("Active").Value) = 0 Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeConcepts&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&Tab=3&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar empleado"" BORDER=""0"" /></A>"
								Else
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeConcepts&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&Tab=3&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar empleado"" BORDER=""0"" /></A>"
								End If
							End If
						sRowContents = sRowContents & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						dTempTotal = 0
						Select Case CInt(oRecordset.Fields("ConceptQttyID").Value)
							Case 1
								dTempTotal = CDbl(oRecordset.Fields("ConceptAmount").Value)
								sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True)
							Case 2
								lErrorNumber = GetConceptAmount(oADODBConnection, aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), aEmployeeComponent(N_ID_EMPLOYEE), aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE), aEmployeeComponent(N_GEOGRAPHICAL_ZONE_ID_EMPLOYEE), aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE), CStr(oRecordset.Fields("AppliesToID").Value), True, False, dTempTotal, sErrorDescription)
								dTempTotal = dTempTotal * CDbl(oRecordset.Fields("ConceptAmount").Value) / 100
								sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True) & "<BR />"
								Call GetNameFromTable(oADODBConnection, "FullConcepts", CStr(oRecordset.Fields("AppliesToID").Value), "<BR />", "", sNames, sErrorDescription)
								sRowContents = sRowContents & "(" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value) & " sobre " & sNames) & ")"
							Case Else
								lErrorNumber = GetConceptAmount(oADODBConnection, aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE), aEmployeeComponent(N_ID_EMPLOYEE), aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE), aEmployeeComponent(N_GEOGRAPHICAL_ZONE_ID_EMPLOYEE), aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE), CStr(oRecordset.Fields("ConceptID").Value), False, False, dTempTotal, sErrorDescription)
								sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True) & "<BR />"
								sRowContents = sRowContents & "(" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value)) & ")"
						End Select
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
								dTotal = dTotal - dTempTotal
							Else
								dTotal = dTotal + dTempTotal
							End If
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CDbl(oRecordset.Fields("ConceptMin").Value) > 0 Then
							If CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 1 Then
								sRowContents = sRowContents & "$" & FormatNumber(CDbl(oRecordset.Fields("ConceptMin").Value), 2, True, False, True)
							Else
								sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptMin").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("MinQttyName").Value))
							End If
						Else
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CDbl(oRecordset.Fields("ConceptMax").Value) > 0 Then
							If CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 1 Then
								sRowContents = sRowContents & "$" & FormatNumber(CDbl(oRecordset.Fields("ConceptMax").Value), 2, True, False, True)
							Else
								sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptMax").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("MaxQttyName").Value))
							End If
						Else
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceTypeName").Value)) & sBoldEnd & sFontEnd

					If CInt(oRecordset.Fields("Active").Value) = 1 Then
						sErrorDescription = "No se pudo obtener la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & ") Order By RecordDate, RecordID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
						If lErrorNumber = 0 Then
							If oPayrollRecordset.EOF Then
								bDisplayForm = True
								sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmYellow.gif"), "ALT=""XXX""", "ALT=""Este concepto no se ha registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>!</B>")
							ElseIf StrComp(FormatNumber(CDbl(oPayrollRecordset.Fields("ConceptAmount").Value), 2, True, False, True), FormatNumber(dTempTotal, 2, True, False, True), vbBinaryCompare) <> 0 Then
								bDisplayForm = True
								sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmRed.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto difiere de lo registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>X</B>")
							Else
								sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmGreen.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto coincide con lo registrado la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>OK</B>")
							End If
							oPayrollRecordset.Close
						Else
							bDisplayForm = True
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					If CInt(oRecordset.Fields("Active").Value) = 1 Then
						sConceptsIDs = sConceptsIDs & CLng(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR
						sConceptsAmounts = sConceptsAmounts & dTempTotal & LIST_SEPARATOR
					End If

					oRecordset.MoveNext
					If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If

			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Credits.*, ConceptShortName, ConceptName, QttyValues.QttyName From Credits, Concepts, QttyValues Where (Credits.CreditTypeID=Concepts.ConceptID) And (Credits.QttyID=QttyValues.QttyID) And (Concepts.ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & ")) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (Credits.FinishDate>=" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ") And (Concepts.EndDate=30000000) Order By OrderInList, ConceptShortName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "," & CStr(oRecordset.Fields("CreditTypeID").Value)
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					sRowContents = ""
					If bForExport Then
						sRowContents = "IcnAlarmXXX.txt"
					Else
						sRowContents = "<IMG SRC=""Images/IcnAlarmXXX.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""XXX"" /><BR />"
					End If
					sRowContents = sRowContents & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
						dTempTotal = 0
						Select Case CInt(oRecordset.Fields("QttyID").Value)
							Case 1
								dTempTotal = CDbl(oRecordset.Fields("PaymentAmount").Value)
								sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True)
							Case 2
								sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus importes."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID In (" & CStr(oRecordset.Fields("AppliesToID").Value) & "))", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
								If lErrorNumber = 0 Then
									If Not oPayrollRecordset.EOF Then
										If Not IsNull(oPayrollRecordset.Fields("TotalAmount").Value) Then dTempTotal = CDbl(oPayrollRecordset.Fields("TotalAmount").Value) * CDbl(oRecordset.Fields("PaymentAmount").Value) / 100
									End If
									oPayrollRecordset.Close
								End If

								sRowContents = sRowContents & "$" & FormatNumber(dTempTotal, 2, True, False, True) & "<BR />"
								Call GetNameFromTable(oADODBConnection, "FullConcepts", CStr(oRecordset.Fields("AppliesToID").Value), "<BR />", "", sNames, sErrorDescription)
								sRowContents = sRowContents & "(" & FormatNumber(CDbl(oRecordset.Fields("PaymentAmount").Value), 2, True, False, True) & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value) & " sobre " & sNames) & ")"
						End Select
						dTotal = dTotal - dTempTotal
					sRowContents = sRowContents & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "<CENTER>---</CENTER>" & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "<CENTER>---</CENTER>" & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "<CENTER>---</CENTER>" & sFontEnd

					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & CStr(oRecordset.Fields("CreditTypeID").Value) & ") Order By RecordDate, RecordID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oPayrollRecordset)
					If lErrorNumber = 0 Then
						If oPayrollRecordset.EOF Then
							bDisplayForm = True
							sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmYellow.gif"), "ALT=""XXX""", "ALT=""Este concepto no se ha registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>!</B>")
						ElseIf StrComp(FormatNumber(CDbl(oPayrollRecordset.Fields("ConceptAmount").Value), 2, True, False, True), FormatNumber(dTempTotal, 2, True, False, True), vbBinaryCompare) <> 0 Then
							bDisplayForm = True
							sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmRed.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto difiere de lo registrado en la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>X</B>")
						Else
							sRowContents = Replace(Replace(Replace(sRowContents, "IcnAlarmXXX.gif", "IcnAlarmGreen.gif"), "ALT=""XXX""", "ALT=""El importe de este concepto coincide con lo registrado la nómina del empleado"""), "IcnAlarmXXX.txt", "<B>OK</B>")
						End If
						oPayrollRecordset.Close
					Else
						bDisplayForm = True
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					sConceptsIDs = sConceptsIDs & CLng(oRecordset.Fields("CreditTypeID").Value) & LIST_SEPARATOR
					sConceptsAmounts = sConceptsAmounts & dTempTotal & LIST_SEPARATOR

					oRecordset.MoveNext
					If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If

			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID In (" & sConceptsToDisable & "))", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "))", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptAmount From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "))", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'			If lErrorNumber = 0 Then
'				If Not oRecordset.EOF Then
'					Do While Not oRecordset.EOF
'						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, StartUserID, EndUserID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & CStr(oRecordset.Fields("ConceptID").Value) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 30000000, " & CStr(oRecordset.Fields("ConceptAmount").Value) & ", 0, 1, 1, 0, 1, 0, 1, -1, 1, 1, 1, -1, -1)", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
'						oRecordset.MoveNext
'						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
'					Loop
'					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
'						Response.Write "window.location.replace('Employees.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Change=1&Tab=3');" & vbNewLine
'					Response.Write "//--></SCRIPT>" & vbNewLine
'				End If
'				oRecordset.Close
'			End If

			sRowContents = TABLE_SEPARATOR & "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>$" & FormatNumber(dTotal, 2, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		Response.Write "</TABLE></DIV>" & vbNewLine

		If Not bForExport Then
			Response.Write "<BR /><BR /><BR />"
			If bDisplayForm Then
				If Len(sConceptsIDs) > 0 Then sConceptsIDs = Left(sConceptsIDs, (Len(sConceptsIDs) - Len(LIST_SEPARATOR)))
				If Len(sConceptsAmounts) > 0 Then sConceptsAmounts = Left(sConceptsAmounts, (Len(sConceptsAmounts) - Len(LIST_SEPARATOR)))
				Response.Write "<FORM NAME=""EmployeeConceptsFrm"" ID=""EmployeeConceptsFrm"" ACTION=""Employees.asp"" METHOD=""GET"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeePayroll"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptsIDs"" ID=""ConceptsIDsHdn"" VALUE=""" & sConceptsIDs & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptsAmounts"" ID=""ConceptsAmountsHdn"" VALUE=""" & sConceptsAmounts & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
					If bDisplayForm And (Len(oRequest("Add").Item) > 0) And (Len(oRequest("SecondUpdate").Item) = 0) And (StrComp(oRequest("Action").Item, "EmployeePayroll", vbBinaryCompare) = 0) Then
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SecondUpdate"" ID=""SecondUpdateHdn"" VALUE=""1"" />"
					End If

					Response.Write "<TABLE BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""1""><TR><TD>"
						Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""5"">"
							Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><B>ICONOGRAFÍA</B></FONT></TD></TR>"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP""><IMG SRC=""Images/IcnAlarmRed.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" /></TD>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">El importe de este concepto difiere de lo registrado en la nómina del empleado.</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP""><IMG SRC=""Images/IcnAlarmYellow.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" /></TD>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Este concepto no se ha registrado en la nómina del empleado.</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP""><IMG SRC=""Images/IcnAlarmGreen.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" /></TD>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">El importe de este concepto coincide con lo registrado la nómina del empleado.</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR><TD COLSPAN=""2"">"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><B>Actualice la nómina del empleado:</B></FONT>"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""130"" HEIGHT=""1"" />"
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Enviar Conceptos"" CLASS=""Buttons"" />"
							Response.Write "</TD></TR>"
						Response.Write "</TABLE>"
					Response.Write "</TD></TR></TABLE>"
				Response.Write "</FORM>"
				If bDisplayForm And (Len(oRequest("Add").Item) > 0) And (Len(oRequest("SecondUpdate").Item) = 0) And (StrComp(oRequest("Action").Item, "EmployeePayroll", vbBinaryCompare) = 0) Then
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "document.EmployeeConceptsFrm.Add.click();" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
			Else
				Call DisplayInstructionsMessage("Confirmación", "Los importes de los conceptos registrados coinciden con los importes registrados en la nómina del empleado.<BR /><BR />")
			End If
		End If
	End If

Call DisplayTimeStamp("FIN")

	Set oPayrollRecordset = Nothing
	Set oRecordset = Nothing
	DisplayEmployeeConceptsTableSp = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeCreditsTable(oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the payment concepts for the given employee
'Inputs:  oRequest, oADODBConnection, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeCreditsTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim bFirst
	Dim lErrorNumber

	lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CreditID, CreditTypeName, StartDate, EndDate, PaymentAmount From Credits, CreditTypes Where (Credits.CreditTypeID=CreditTypes.CreditTypeID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By StartDate, CreditTypeName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
					If bForExport Then
						Response.Write "1"
					Else
						Response.Write "0"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("Crédito,Fecha de inicio,Fecha final,Importe", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,RIGHT", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If

					bFirst = True
					Do While Not oRecordset.EOF
						sFontBegin = ""
						sFontEnd = ""
						sRowContents = "<INPUT TYPE=""RADIO"" NAME=""CreditID"" ID=""CreditIDRd"" VALUE=""" & CStr(oRecordset.Fields("CreditID").Value) & """"
							If bFirst Then sRowContents = sRowContents & " CHECKED=""1"""
						sRowContents = sRowContents & " />"
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("CreditTypeName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CDbl(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CDbl(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & FormatNumber(CDbl(oRecordset.Fields("PaymentAmount").Value), 2, True, False, True) & sFontEnd
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						bFirst = False
						oRecordset.MoveNext
						If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
				Response.Write "</TABLE></DIV>"
			Else
				lErrorNumber = -1
				sErrorDescription = "El empleado no tiene descuentos registrados en el sistema."
			End If
			oRecordset.Close
		End If
	End If
	aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) = "-2,-1,0,55," & MAIN_CONCEPTS_FOR_PAYROLL

	Set oPayrollRecordset = Nothing
	Set oRecordset = Nothing
	DisplayEmployeeCreditsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesDocumentsTable(oRequest, oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the documents for service sheet
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesDocumentsTable"
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
	Dim oStartDate
	Dim lDate
	Dim sAuthorizers
	Dim asAuthorizers
	Dim bAuthorize
	Dim iIndex
	Dim bPrinted
	Dim bHasAuthorisations
	Dim bAuthorizedForMe
	Dim bNeedMyAuthorization

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetEmployeesDocuments(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Then
					asColumnsTitles = Split("N. de empleado,Nombre del empleado,Número de documento,Fecha de solicitud,Usuario que capturó,T. Documento", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,300,200,200,200,100",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,CENTER,,,CENTER", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Acciones,N. de empleado,Nombre del empleado,Número de documento,Fecha de solicitud,Usuario que capturó,T. Documento", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,100,300,200,200,200,100",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,,CENTER,,,CENTER", ",", -1, vbBinaryCompare)
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
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(oRecordset.Fields("EmployeeID").Value), oRequest("EmployeeID").Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields("StartDate").Value), oRequest("StartDate").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					sRowContents = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1 Then
							If False Then
								If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID &""">"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("AccountID").Value) & """ ID=""" & CStr(oRecordset.Fields("AccountID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("AccountID").Value) & """ CHECKED=""1"" />"
									End If
								End If
							End If
						Else
							bPrinted = (CInt(oRecordset.Fields("bPrinted").Value) = 1)
							bHasAuthorisations = (StrComp(CStr(oRecordset.Fields("Authorized").Value), "-1", vbBinaryCompare) <> 0)
							bAuthorizedForMe = (InStr(1, "," & oRecordset.Fields("Authorized").Value & ",", "," & CStr(aLoginComponent(N_USER_ID_LOGIN)) & ",", vbBinaryCompare) > 0)
							bNeedMyAuthorization = (InStr(1, "," & oRecordset.Fields("Authorizers").Value & ",", "," & CStr(aLoginComponent(N_USER_ID_LOGIN)) & ",", vbBinaryCompare) > 0)
							If Not bPrinted Then
								bAuthorize = True
								sAuthorizers = CStr(oRecordset.Fields("Authorizers").Value)
								asAuthorizers = Split(sAuthorizers, ",")
								For iIndex = 0 To UBound(asAuthorizers)
									If InStr(1, "," & oRecordset.Fields("Authorized").Value & ",", "," & CStr(asAuthorizers(iIndex)) & ",", vbBinaryCompare) = 0 Then
										bAuthorize = False
										Exit For
									End If
								Next
								' Eliminar solicitud
								If (Not bHasAuthorisations) And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&SectionID=" & iSectionID & "&Remove=1&DocumentDate=" & CStr(oRecordset.Fields("DocumentDate").Value) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								' Modificar solicitud
								If (Not bHasAuthorisations) And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&SectionID=" & iSectionID & "&ServiceSheetChange=1&DocumentDate=" & CStr(oRecordset.Fields("DocumentDate").Value) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar solicitud"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								' Autorizar solicitud
								If bNeedMyAuthorization And (Not bAuthorizedForMe) And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&SectionID=" & iSectionID & "&Authorize=1&DocumentDate=" & CStr(oRecordset.Fields("DocumentDate").Value) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Autorizo solicitud"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								'Mostrar detalle en pantalla
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&SectionID=" & iSectionID & "&ShowServiceSheet=1&DocumentDate=" & CStr(oRecordset.Fields("DocumentDate").Value) & "&DocumentTypeID=" & CInt(oRecordset.Fields("DocumentTypeID").Value) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnForm.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Visualizar previo"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								'Generar docuumentos
								If bAuthorize And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Main_ISSSTE.asp" & "?Action=" & sAction & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&SectionID=" & iSectionID & "&GenerateReport=1&DocumentDate=" & CStr(oRecordset.Fields("DocumentDate").Value) & "&DocumentTypeID=" & CInt(oRecordset.Fields("DocumentTypeID").Value) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Generar documentos"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							Else
								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "<A HREF=""" & REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_1203_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & CStr(oRecordset.Fields("ReportName").Value) & ".zip"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Descargar documentos"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
						End If
						sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
					End If
					If bForExport Then
						sRowContents = sRowContents & "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
					Else
						sRowContents = sRowContents & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("DocumentDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value))
					If CInt(oRecordset.Fields("DocumentTypeID").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Completo")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Normal")
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				sErrorDescription = "Introduzca un numero de empleado para consultar las solicitudes de Hojas únicas de servicio registradas."
			Else
				sErrorDescription = "No se han registrado solicitudes de Hojas únicas de servicio para este empleado."
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesDocumentsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeHistoryList(oRequest, oADODBConnection, bForExport, bFull, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the history list for the employee from
'         the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeHistoryList"
	Dim oRecordset
	Dim sCondition
	Dim sBoldBegin
	Dim sBoldEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim oRecordsetConcept

	Dim sEmployeeTypeShortName
	Dim sEmployeeTypeName
	Dim sCompanyShortName
	Dim sCompanyName
	Dim sServiceShortName
	Dim sServiceName
	Dim sLevelShortName
	Dim sLevelName
	Dim lPeriod
	Dim sCatalogShortName
	Dim sCatalogName
	Dim oRecordsetCatalog
	Dim lCurrPayrollID

	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "EmployeeDate", False, sCondition)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EHL.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	If (InStr(1, ",-75,-64,1,2,3,4,5,6,8,10,12,13,14,17,18,21,26,28,29,30,31,32,33,34,37,38,39,40,41,43,44,45,46,47,48,50,51,53,62,63,66,68,", "," & oRequest("ReasonID").Item & ",", vbBinaryCompare) = 0) Then
		If False Then
			sCondition = sCondition & " And (((Zones.StartDate>=EHL.EmployeeDate) And (Zones.StartDate<=EHL.EndDate)) Or ((Zones.EndDate>=EHL.EmployeeDate) And (Zones.EndDate<=EHL.EndDate)) Or ((Zones.EndDate>=EHL.EmployeeDate) And (Zones.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Areas.StartDate>=EHL.EmployeeDate) And (Areas.StartDate<=EHL.EndDate)) Or ((Areas.EndDate>=EHL.EmployeeDate) And (Areas.EndDate<=EHL.EndDate)) Or ((Areas.EndDate>=EHL.EmployeeDate) And (Areas.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((PaymentCenters.StartDate>=EHL.EmployeeDate) And (PaymentCenters.StartDate<=EHL.EndDate)) Or ((PaymentCenters.EndDate>=EHL.EmployeeDate) And (PaymentCenters.EndDate<=EHL.EndDate)) Or ((PaymentCenters.EndDate>=EHL.EmployeeDate) And (PaymentCenters.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Positions.StartDate>=EHL.EmployeeDate) And (Positions.StartDate<=EHL.EndDate)) Or ((Positions.EndDate>=EHL.EmployeeDate) And (Positions.EndDate<=EHL.EndDate)) Or ((Positions.EndDate>=EHL.EmployeeDate) And (Positions.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((PositionTypes.StartDate>=EHL.EmployeeDate) And (PositionTypes.StartDate<=EHL.EndDate)) Or ((PositionTypes.EndDate>=EHL.EmployeeDate) And (PositionTypes.EndDate<=EHL.EndDate)) Or ((PositionTypes.EndDate>=EHL.EmployeeDate) And (PositionTypes.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Shifts.StartDate>=EHL.EmployeeDate) And (Shifts.StartDate<=EHL.EndDate)) Or ((Shifts.EndDate>=EHL.EmployeeDate) And (Shifts.EndDate<=EHL.EndDate)) Or ((Shifts.EndDate>=EHL.EmployeeDate) And (Shifts.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Journeys.StartDate>=EHL.EmployeeDate) And (Journeys.StartDate<=EHL.EndDate)) Or ((Journeys.EndDate>=EHL.EmployeeDate) And (Journeys.EndDate<=EHL.EndDate)) Or ((Journeys.EndDate>=EHL.EmployeeDate) And (Journeys.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=EHL.EmployeeDate) And (GroupGradeLevels.StartDate<=EHL.EndDate)) Or ((GroupGradeLevels.EndDate>=EHL.EmployeeDate) And (GroupGradeLevels.EndDate<=EHL.EndDate)) Or ((GroupGradeLevels.EndDate>=EHL.EmployeeDate) And (GroupGradeLevels.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Companies.StartDate>=EHL.EmployeeDate) And (Companies.StartDate<=EHL.EndDate)) Or ((Companies.EndDate>=EHL.EmployeeDate) And (Companies.EndDate<=EHL.EndDate)) Or ((Companies.EndDate>=EHL.EmployeeDate) And (Companies.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((EmployeeTypes.StartDate>=EHL.EmployeeDate) And (EmployeeTypes.StartDate<=EHL.EndDate)) Or ((EmployeeTypes.EndDate>=EHL.EmployeeDate) And (EmployeeTypes.EndDate<=EHL.EndDate)) Or ((EmployeeTypes.EndDate>=EHL.EmployeeDate) And (EmployeeTypes.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Services.StartDate>=EHL.EmployeeDate) And (Services.StartDate<=EHL.EndDate)) Or ((Services.EndDate>=EHL.EmployeeDate) And (Services.EndDate<=EHL.EndDate)) Or ((Services.EndDate>=EHL.EmployeeDate) And (Services.StartDate<=EHL.EndDate)))"
			sCondition = sCondition & " And (((Levels.StartDate>=EHL.EmployeeDate) And (Levels.StartDate<=EHL.EndDate)) Or ((Levels.EndDate>=EHL.EmployeeDate) And (Levels.EndDate<=EHL.EndDate)) Or ((Levels.EndDate>=EHL.EmployeeDate) And (Levels.StartDate<=EHL.EndDate)))"
		End If
	End If
	sQuery = "Select PayrollID From Payrolls Where (IsClosed = 0) And (PayrollTypeID = 1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery , "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lCurrPayrollID = oRecordset.Fields("PayrollID").Value
	If (InStr(1, ",-75,-64,1,2,3,4,5,6,8,10,12,13,14,17,18,21,26,28,29,30,31,32,33,34,37,38,39,40,41,43,44,45,46,47,48,50,51,53,62,63,66,68,", "," & oRequest("ReasonID").Item & ",", vbBinaryCompare) <> 0) And _
		 (InStr(1,oRequest.Item,"Tab",vbBinaryCompare) = 0) Then
		'sQuery = "Select Distinct EmployeesHistoryList.EmployeeID, EmployeesHistoryList.ReasonID, EmployeesHistoryList.EmployeeNumber, CompanyShortName, CompanyName, EmployeeTypeShortName, EmployeeTypeName, EmployeesHistoryList.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ServiceShortName, ServiceName, LevelShortName, EmployeesHistoryList.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EmployeesHistoryList.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EmployeesHistoryList.WorkingHours, StatusName, EmployeesHistoryList.ReasonID, ReasonName, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate From EmployeesHistoryList, Companies, EmployeeTypes, Zones, Areas, Areas As PaymentCenters, Positions, PositionTypes, Shifts, Journeys, GroupGradeLevels, Services, Levels, StatusEmployees, Reasons Where (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & sCondition & " And (StatusEmployees.StatusReasonID = 0) Order By EmployeesHistoryList.EmployeeID, EmployeeDate Desc, EmployeesHistoryList.EndDate Desc"
		sQuery = "Select Distinct EHL.EmployeeID, EHL.EmployeeNumber, EHL.CompanyID, CompanyShortName, CompanyName, EHL.EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName, EHL.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EHL.ServiceID, ServiceShortName, ServiceName, EHL.LevelID, LevelShortName, EHL.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EHL.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EHL.WorkingHours, StatusName, EHL.ReasonID, ReasonName, EHL.EmployeeDate, EHL.EndDate From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, PositionTypes, Shifts, Journeys, GroupGradeLevels, StatusEmployees, Reasons, Companies, EmployeeTypes, Services, Levels Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.ServiceID=Services.ServiceID) And (EHL.LevelID=Levels.LevelID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.ShiftID = Shifts.ShiftID) And (EHL.JourneyID = Journeys.JourneyID) And (EHL.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (EHL.PositionID=Positions.PositionID) And (EHL.PositionTypeID=PositionTypes.PositionTypeID) And (EHL.ShiftID=Shifts.ShiftID) And (EHL.JourneyID=Journeys.JourneyID) And (EHL.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & sCondition & " And (StatusEmployees.StatusReasonID = 0) And (EHL.ReasonID <> 0) Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc"
	Else
		'sQuery = "Select Distinct EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, CompanyShortName, CompanyName, EmployeeTypeShortName, EmployeeTypeName, EmployeesHistoryList.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ServiceShortName, ServiceName, LevelShortName, EmployeesHistoryList.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EmployeesHistoryList.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EmployeesHistoryList.WorkingHours, StatusName, EmployeesHistoryList.ReasonID, ReasonName, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate From EmployeesHistoryList, Companies, EmployeeTypes, Zones, Areas, Areas As PaymentCenters, Positions, PositionTypes, Shifts, Journeys, GroupGradeLevels, Services, Levels, StatusEmployees, Reasons Where (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeeDate Desc, EmployeesHistoryList.EndDate Desc"
		sQuery = "Select Distinct EHL.EmployeeID, EHL.EmployeeNumber, EHL.CompanyID, CompanyShortName, CompanyName, EHL.EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName, EHL.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EHL.ServiceID, ServiceShortName, ServiceName, EHL.LevelID, LevelShortName, EHL.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EHL.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EHL.WorkingHours, StatusName, EHL.ReasonID, ReasonName, EHL.EmployeeDate, EHL.EndDate, EHL.UserID, EHL.PayrollDate, UserName, UserLastName, COUNT(*) From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, PositionTypes, Shifts, Journeys, GroupGradeLevels, StatusEmployees, Reasons, Companies, EmployeeTypes, Services, Levels, Users Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.ServiceID=Services.ServiceID) And (EHL.LevelID=Levels.LevelID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.ShiftID = Shifts.ShiftID) And (EHL.JourneyID = Journeys.JourneyID) And (EHL.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (EHL.PositionID=Positions.PositionID) And (EHL.PositionTypeID=PositionTypes.PositionTypeID) And (EHL.ShiftID=Shifts.ShiftID) And (EHL.JourneyID=Journeys.JourneyID) And (EHL.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.UserId = Users.UserID) And (EHL.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & sCondition & " And (EHL.ReasonID <> 0) Group by EHL.EmployeeID, EHL.EmployeeNumber, EHL.CompanyID, CompanyShortName, CompanyName, EHL.EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName, EHL.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode, PaymentCenters.AreaName, EHL.ServiceID, ServiceShortName, ServiceName, EHL.LevelID, LevelShortName, EHL.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EHL.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EHL.WorkingHours, StatusName, EHL.ReasonID, ReasonName, EHL.EmployeeDate, EHL.EndDate, EHL.UserID, EHL.PayrollDate, UserName, UserLastName Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc"
	End If
	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery , "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If bFull Then
				Response.Write "<TABLE WIDTH=""3000"" BORDER="""
			Else
				Response.Write "<TABLE BORDER="""
			End If
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bFull Then
					If InStr(1,oRequest.Item,"Tab",vbBinaryCompare) > 0 Then
						asColumnsTitles = "Quincena, Usuario, Fecha inicio,Fecha fin,Compañía,Tipo de tabulador,Plaza,Tipo de movimiento,Estatus,Puesto,Tipo de puesto,Adscripción,Centro de pago,Servicio,Nivel-subnivel,Clasificación,Grupo grado nivel,Integración,Turno,Horario,Horas laboradas"
						asCellWidths = "100,50,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100"
						asCellAlignments = ",,,,,,,,CENTER,,,,,,CENTER,CENTER,,CENTER,,,CENTER"
					Else
						asColumnsTitles = "Fecha inicio,Fecha fin,Compañía,Tipo de tabulador,Plaza,Tipo de movimiento,Estatus,Puesto,Tipo de puesto,Adscripción,Centro de pago,Servicio,Nivel-subnivel,Clasificación,Grupo grado nivel,Integración,Turno,Horario,Horas laboradas"
						asCellWidths = "100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100"
						asCellAlignments = ",,,,,,CENTER,,,,,,CENTER,CENTER,,CENTER,,,CENTER"
					End If
					If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CertificacionesYArchivo & ",", vbBinaryCompare) > 0) And (Not bForExport) Then
						asColumnsTitles = "Acciones," & asColumnsTitles
						asCellWidths = "80," & asCellWidths
						asCellAlignments = "CENTER," & asCellAlignments
					End If
				Else
					asColumnsTitles = "Fecha inicio,Fecha fin,Plaza,Tipo de movimiento"
					asCellWidths = "100,100,100,100,100"
					asCellAlignments = ",,,,"
				End If
				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				aEmployeeComponent(N_HISTORY_LIST_RECORTS) = 0
				Do While Not oRecordset.EOF
					lPeriod = oRecordset.Fields("EmployeeDate").Value & ":" & oRecordset.Fields("EndDate").Value
					Call GetEmployeeDataFromCatalog(oADODBConnection, "EmployeeTypes", "EmployeeTypeShortName, EmployeeTypeName", "EmployeeTypeID", oRecordset.Fields("EmployeeTypeID").Value, lPeriod, oRecordsetCatalog, sErrorDescription)
					sEmployeeTypeShortName = sCatalogShortName
					sEmployeeTypeName = sCatalogName
					Call GetEmployeeDataFromCatalog(oADODBConnection, "Companies", "CompanyShortName, CompanyName", "CompanyID", oRecordset.Fields("CompanyID").Value, lPeriod, oRecordsetCatalog, sErrorDescription)
					sCompanyShortName = sCatalogShortName
					sCompanyName = sCatalogName
					Call GetEmployeeDataFromCatalog(oADODBConnection, "Services", "ServiceShortName, ServiceName", "ServiceID", oRecordset.Fields("ServiceID").Value, lPeriod, oRecordsetCatalog, sErrorDescription)
					sServiceShortName = sCatalogShortName
					sServiceName = sCatalogName
					Call GetEmployeeDataFromCatalog(oADODBConnection, "Levels", "LevelShortName, LevelName", "LevelID", oRecordset.Fields("CompanyID").Value, lPeriod, oRecordsetCatalog, sErrorDescription)
					sLevelShortName = sCatalogShortName
					sLevelName = sCatalogName

					aEmployeeComponent(N_HISTORY_LIST_RECORTS) = aEmployeeComponent(N_HISTORY_LIST_RECORTS) + 1
					sBoldBegin = ""
					sBoldEnd = ""
					If (CLng(oRequest("EmployeeDate").Item) = CLng(oRecordset.Fields("EmployeeDate").Value)) Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If bFull Then
						If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CertificacionesYArchivo & ",", vbBinaryCompare) > 0) And (Not bForExport) Then
							If aLoginComponent(N_PROFILE_ID_LOGIN) <> 1 Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&PayrollDate=" & oRecordset.Fields("PayrollDate").Value & "&Tab=6&Change=1&ReportID=707&" & aEmployeeComponent(S_URL_EMPLOYEE) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=6&Delete=1&ReportID=707&" & aEmployeeComponent(S_URL_EMPLOYEE) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								sRowContents = sRowContents & TABLE_SEPARATOR
							Else
								If CLng(oRecordset.Fields("PayrollDate").Value) = CLng(lCurrPayrollID) Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=6&Change=1&ReportID=707&" & aEmployeeComponent(S_URL_EMPLOYEE) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=6&Delete=1&ReportID=707&" & aEmployeeComponent(S_URL_EMPLOYEE) & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR
								Else
									sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
								End If
							End If
						End If
					End If
					If InStr(1,oRequest.Item,"Tab",vbBinaryCompare) > 0 Then
						If CLng(oRecordset.Fields("PayrollDate").Value) <= 0 Then
							sRowContents = sRowContents & "No Disponible"
						Else
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollDate").Value), -1, -1, -1)
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & oRecordset.Fields("UserLastName").Value & ", " & oRecordset.Fields("UserName").Value & TABLE_SEPARATOR
					End If
					If CLng(oRecordset.Fields("EmployeeDate").Value) = 0 Then
						sRowContents = sRowContents & sBoldBegin & "-" & sBoldEnd
					Else
						sRowContents = sRowContents & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value), -1, -1, -1) & sBoldEnd
					End If
					If (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Indefinida" & sBoldEnd
					ElseIf CLng(oRecordset.Fields("EndDate")) < 100 Then
						'sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value), -1, -1, -1) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "---" & sBoldEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd
					End If
					If bFull Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value)) & sBoldEnd
					End If
					If (CLng(oRecordset.Fields("JobID").Value) = -2) Or (CLng(oRecordset.Fields("JobID").Value) = -3) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "No Definida" & sBoldEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)) & sBoldEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR
					sCondition = ""
					If (CLng(oRecordset.Fields("JobID").Value) = -2) Then
						sRowContents = sRowContents & sBoldBegin & "Alta" & sBoldEnd
					ElseIf (CLng(oRecordset.Fields("JobID").Value) = -3) Then
						sRowContents = sRowContents & sBoldBegin & "Licencia sin goce de sueldo" & sBoldEnd
					Else
						sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)) & sBoldEnd
					End If
					If bFull Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value) & ". " & CStr(oRecordset.Fields("PositionTypeName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & sBoldEnd
						If CStr(oRecordset.Fields("ClassificationID").Value) <> "-1" Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value)) & sBoldEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Ninguno" & sBoldEnd
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ". " & CStr(oRecordset.Fields("GroupGradeLevelName").Value)) & sBoldEnd
						If CStr(oRecordset.Fields("IntegrationID").Value) <> "-1" Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value)) & sBoldEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Ninguno" & sBoldEnd
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & sBoldEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & sBoldEnd
					End If
					Err.Clear
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = "-1"
			sErrorDescription = "No existen registros que cumplan con el criterio de la búsqueda"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeeHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeePayroll(oRequest, oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the given employee
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeePayroll"
	Dim oRecordset
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, OrderInList, ConceptAmount, RecordDate, PayRollTypeName From Payroll_" & oRequest("PayrollID").Item & ", Concepts, PayRollTypes Where (Payroll_" & oRequest("PayrollID").Item & ".ConceptID=Concepts.ConceptID) And (Payroll_" & oRequest("PayrollID").Item & ".PayRollTypeID=PayRollTypes.PayRollTypeID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (Concepts.StartDate<=" & oRequest("PayrollID").Item & ") And (Concepts.EndDate>=" & oRequest("PayrollID").Item & ") And (ConceptAmount<>0) Order By IsDeduction, OrderInList, ConceptShortName, ConceptName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Concepto,Importe,Fecha de registro,Tipo de nómina", ",", -1, vbBinaryCompare)
				asCellWidths = Split("300,200,200,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",RIGHT,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					sBoldBegin = ""
					sBoldEnd = ""
					If (CInt(oRecordset.Fields("IsDeduction").Value) = 1) Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					If InStr(1, ",-2,-1,0,", ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
						Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
					End If
					sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "$" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("RecordDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PayRollTypeName").Value)) & sBoldEnd & sFontEnd

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					If InStr(1, ",-1,0,", ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0 Then
						Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
					End If

					oRecordset.MoveNext
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen pagos para este empleado en el periodo indicado."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeePayroll = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesBanksAccountsTable(oRequest, oADODBConnection, lIDColumn, bForExport, lStartPage, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesBanksAccountsTable"
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
	Dim oStartDate
	Dim lDate
	Dim sAccount
	Dim sSucursal
	Dim sAccountNumber

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetEmployeesBankAccounts(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetchForSections(oRequest, lStartPage, ROWS_REPORT, aEmployeeComponent(N_ACTIVE_EMPLOYEE), oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Then
					asColumnsTitles = Split("N. de empleado,Nombre del empleado,Número de cuenta,Banco,Fecha de inicio,Fecha de fin, Usuario que capturó", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,200,300,200,200,400",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,CENTER,,,,CENTER", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Acciones,N. de empleado,Nombre del empleado,Número de cuenta,Banco,Fecha de inicio,Fecha de fin, Usuario que capturó", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,500,200,300,200,200,400",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,,CENTER,,,,CENTER", ",", -1, vbBinaryCompare)
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
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("AccountID").Value), oRequest("AccountID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Removed").Value) = 1 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					sSucursal = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1 Then
							If False Then
								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							If False And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&BankAccountChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If False Then
								If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID &""">"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("AccountID").Value) & """ ID=""" & CStr(oRecordset.Fields("AccountID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("AccountID").Value) & """ CHECKED=""1"" />"
									End If
								End If
							End If
						Else
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("AccountID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
					End If
					If bForExport Then
						sRowContents = sRowContents & "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
					Else
						sRowContents = sRowContents & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
					If (InStr(1, CStr(oRecordset.Fields("AccountNumber").Value), ".", vbBinaryCompare) = 0) Then
						sAccountNumber = Split(CStr(oRecordset.Fields("AccountNumber").Value), LIST_SEPARATOR)
						sAccount = sAccountNumber(0)
						sSucursal = sAccountNumber(1)
						If Len(sSucursal) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sAccount)
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sAccount & "-" & sSucursal)
						End If
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Cheques")
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sErrorDescription = "Introduzca un numero de empleado para consultar las cuentas bancarias registradas."
				Else
					sErrorDescription = "No existen cuentas bancarias en proceso para ser aplicadas para el empleado indicado."
				End If
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sErrorDescription = "No se han registrado cuentas bancarias para este empleado."
				Else
					sErrorDescription = "No se han registrado cuentas bancarias en proceso para este empleado."
				End If
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesBanksAccountsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesGradesTable(oRequest, oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesGradesTable"
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
	Dim oStartDate
	Dim lDate

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetEmployeesGrades(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Then
					asColumnsTitles = Split("N. de empleado,Nombre del empleado,Fecha de inicio,Fecha de termino,Quincena a considerar,Grado,Usuario que capturó", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,200,300,200,200,400",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,CENTER,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Acciones,N. de empleado,Nombre del empleado,Fecha de inicio,Fecha de termino,Quincena a considerar,Grado,Usuario que capturó", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,500,200,300,200,200,400",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,,CENTER,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
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
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(oRecordset.Fields("EmployeeID").Value), oRequest("EmployeeID").Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields("StartDate").Value), oRequest("StartDate").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					sRowContents = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1 Then
							If False Then
								If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID &""">"
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("AccountID").Value) & """ ID=""" & CStr(oRecordset.Fields("AccountID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("AccountID").Value) & """ CHECKED=""1"" />"
									End If
								End If
							End If
						Else
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
					End If
					If bForExport Then
						sRowContents = sRowContents & "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
					Else
						sRowContents = sRowContents & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeGrade").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sErrorDescription = "Introduzca un numero de empleado para consultar las calificaciones registradas."
				Else
					sErrorDescription = "No existen calificaciones en proceso para ser aplicadas del empleado indicado."
				End If
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sErrorDescription = "No se han registrado calificaciones para este empleado."
				Else
					sErrorDescription = "No se han registrado calificaciones en proceso para este empleado."
				End If
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesGradesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the employees from
'         the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesTable"
	Dim sRequest
	Dim iRecordCounter
	Dim oRecordset
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
	Dim sAction
	Dim lErrorNumber

	lErrorNumber = GetEmployees(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				sRequest = RemoveParameterFromURLString(RemoveEmptyParametersFromURLString(oRequest), "ReinsuranceDisaster")
				If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					sRowContents = "Acciones"
					asCellWidths = asCellWidths & "80"
				Else
					sRowContents = ""
					asCellWidths = asCellWidths & "20"
				End If
				asCellAlignments = asCellAlignments & "CENTER"

				sRowContents = sRowContents & TABLE_SEPARATOR & "No. Empleado"
				If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asCellWidths = asCellWidths & ",120"
				Else
					asCellWidths = asCellWidths & ",180"
				End If
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "EmployeeNumber", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeNumber"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeNumber"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeNumber"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Nombre"
				asCellWidths = asCellWidths & ",250"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "EmployeeLastName, EmployeeLastName2, EmployeeName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Plaza"
				asCellWidths = asCellWidths & ",150"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "JobNumber", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "JobNumber"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "JobNumber"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "JobNumber"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Zona"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "ZoneName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "ZoneName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "ZoneName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "ZoneName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Área"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "AreaName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "AreaName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "AreaName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "AreaName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Puesto"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "PositionShortName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "PositionShortName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "PositionShortName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "PositionShortName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Nivel-Subnivel"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "LevelName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "LevelName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "LevelName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "LevelName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				sRowContents = sRowContents & TABLE_SEPARATOR & "Estatus"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE), "StatusName", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If

				asColumnsTitles = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				sAction = "ShowInfo"
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sAction = "Change"
				iRecordCounter = 0
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("EmployeeID").Value), oRequest("EmployeeID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""EmployeeID"" ID=""EmployeeIDRd"" VALUE=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""EmployeeID"" ID=""EmployeeIDChk"" VALUE=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ />"
						Case Else
							If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
								sRowContents = sRowContents & "&nbsp;"
									If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
										sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Tab=1&Change=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									End If

									'If B_DELETE And (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
									'	sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Tab=1&Delete=1"">"
									'		sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									'	sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									'End If

									'If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									'	If CInt(oRecordset.Fields("Active").Value) = 0 Then
									'		sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Tab=1&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar empleado"" BORDER=""0"" /></A>"
									'	Else
									'		sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Tab=1&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar empleado"" BORDER=""0"" /></A>"
									'	End If
									'End If
								sRowContents = sRowContents & "&nbsp;"
							End If
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						If Not bForExport Then sRowContents = sRowContents & " HREF=""Employees.asp?EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Tab=1&" & sAction & "=1"""
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</A>" & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "<A"
						If Not bForExport Then
							If Not IsNull(oRecordset.Fields("EmployeeEmail").Value) Then
								If Len(CStr(oRecordset.Fields("EmployeeEmail").Value)) > 0 Then sRowContents = sRowContents & " HREF=""mailto: " & CStr(oRecordset.Fields("EmployeeEmail").Value) & """"
							End If
						End If
					sRowContents = sRowContents & ">"
						If StrComp(CStr(oRecordset.Fields("EmployeeName").Value), CStr(oRecordset.Fields("EmployeeLastName").Value), vbBinaryCompare) <> 0 Then
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
							End If
						Else
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd & sFontEnd

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
			sErrorDescription = "No existen empleados que cumplan con los criterios de la búsqueda."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayFM1Table(oRequest, oADODBConnection, bForExport, iStatusID, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the employees
'         whose has status = -3 or -4
'Inputs:  oRequest, oADODBConnection, bForExport, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayValidateFM1Table"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
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
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim sLinkMessage1
	Dim sLinkMessage2
	Dim sLinkRemove

	sErrorDescription = "No existen empleados con estatus en proceso de autorización."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.ModifyDate, Employees.EmployeeNumber, Employees.EmployeeTypeID, EmployeeTypes.EmployeeTypeShortName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC From Employees, EmployeesHistoryList, EmployeeTypes Where (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.StatusID=" & iStatusID & ") Order By EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.ModifyDate, Employees.EmployeeID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Select Case iStatusID
			Case -3
				sLinkMessage1 = "Validar"
				sLinkMessage2 = "Rechazar"
				sLinkRemove = "Malformed"
			Case -4
				sLinkMessage1 = "Autorizar"
				sLinkMessage2 = "Rechazar"
				sLinkRemove = "Unauthorized"
		End Select

		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine

			asColumnsTitles = Split("Fecha de registro,Fecha inicio vigencia,Número de empleado,Tipo de Tabulador,Apellido paterno,Apellido materno,Nombre,Acciones", ",", -1, vbBinaryCompare)
			asCellWidths = Split(",,,,,,,", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & " "
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))

				'If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
				'	sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "Employees.asp" & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&DisplayFM1=1"">"
				'	sRowContents = sRowContents & "<IMG SRC=""Images/IcnForm.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Consultar FM1"" BORDER=""0"" />"
				'	sRowContents = sRowContents & "</A>&nbsp;"
				'End If
				If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "Employees.asp" & "?Action=EmployeesNew&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "&DisplayFM1=1"">"
						sRowContents = sRowContents & "<IMG SRC=""Images/IcnForm.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Consultar FM1"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
				If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Employees.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & """>"
						sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=" & sLinkMessage1 & " BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
				If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "Employees.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Modify=" & sLinkRemove & """>"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=" & sLinkMessage2 & " BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
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
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen empleados con estatus en proceso de autorización."
		End If
	End If
	
	Set oRecordset = Nothing
	DisplayValidateFM1Table = lErrorNumber
	Err.Clear
End Function

Function DisplayMedicalAreasTable(oRequest, oADODBConnection, bForExport, lStartPage, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMedicalAreasTable"
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
	Dim oStartDate
	Dim lDate

	oStartDate = Now()
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from MedicalAreas, MedicalAreasTypes, Companies, Positions, Services Where MedicalAreas.CompanyID=Companies.CompanyID And MedicalAreas.MedicalAreasTypeID=MedicalAreasTypes.MedicalAreasTypeID And MedicalAreas.PositionID=Positions.PositionID And MedicalAreas.ServiceID=Services.ServiceID", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, lStartPage, ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Then
					asColumnsTitles = Split("Compañía,Tipo de reporte UNIMED,Puesto,Servicio,No.Anexo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,200,300,200",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,CENTER,,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Acciones,Compañía,Tipo de reporte UNIMED,Puesto,Servicio,No.Anexo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,500,200,300,200",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,,CENTER,,", ",", -1, vbBinaryCompare)
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
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("MedicalAreasID").Value), oRequest("MedicalAreasID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Removed").Value) = 1 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ReasonID=" & lReasonID & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&AccountID=" & CStr(oRecordset.Fields("AccountID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("AccountID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
						End If
						sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
					End If
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("MedicalAreasTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ColumnNumber").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros UNIMED para ser exportados."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayMedicalAreasTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeBeneficiariesTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the employee's
'         beneficiaries from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeBeneficiariesTable"
	Dim sRequest
	Dim iRecordCounter
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sAction
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.*, AlimonyTypeName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName From EmployeesBeneficiariesLKP, AlimonyTypes, Areas As PaymentCenters Where (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By BeneficiaryNumber, BeneficiaryLastName, BeneficiaryLastName2, BeneficiaryName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = "Número,Nombre,Fecha de nacimiento,Importe,Tipo"
				asCellWidths = "70,200,100,100,200"
				asCellAlignments = ",,,,"
				asCellAlignments
				If Not(bForExport) Then
					asColumnsTitles = "Acciones," & asColumnsTitles
					asCellWidths = "100," & asCellWidths
					asCellAlignments = "CENTER," & asCellAlignments
				End If
				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("BeneficiaryID").Value), oRequest("BeneficiaryID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDRd"" VALUE=""" & CStr(oRecordset.Fields("BeneficiaryID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDChk"" VALUE=""" & CStr(oRecordset.Fields("BeneficiaryID").Value) & """ />"
'						Case Else
'							If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
'								sRowContents = sRowContents & "&nbsp;"
'									If (CInt(Request.Cookies("SIAP_SubSectionID")) = 11) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 14) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 17) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 12) Then
'										sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
'									Else
'										If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
'											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & CStr(oRecordset.Fields("BeneficiaryID").Value) & "&Tab=1&Change=1"">"
'												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
'											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
'										End If
'										If B_DELETE And (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
'											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Employees&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & CStr(oRecordset.Fields("BeneficiaryID").Value) & "&Tab=1&Delete=1"">"
'												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
'											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
'										End If
'									End If
'								sRowContents = sRowContents & "&nbsp;"
'							End If
					End Select
					If bForExport Then
						sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) & sBoldEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) & sBoldEnd
					End If

					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryLastName").Value) & " " & CStr(oRecordset.Fields("BeneficiaryLastName2").Value) & ", " & CStr(oRecordset.Fields("BeneficiaryName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin
						If CLng(oRecordset.Fields("BeneficiaryBirthDate").Value) > 0 Then 
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("BeneficiaryBirthDate").Value), -1, -1, -1)
						Else
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						End If
					sRowContents = sRowContents & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), -1, -1, -1) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeName").Value)) & sBoldEnd

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			Response.Write "<BR />- El empleado no tiene beneficiarios registrados.<BR /><BR />"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeeBeneficiariesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesCreditorsTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the beneficiaries of employees
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesCreditorsTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
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
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion
	Dim iBeneficiaryID
	Dim iStartDate
	Dim bIsFirst
	Dim sConceptNames

	bIsFirst = True
	sErrorDescription = "No existen acreedores registrados para el empledado."
	sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," &_
			" CreditorNumber, CreditorName + ' ' + CreditorLastName + ' ' + CreditorLastName2 As CreditorFullName," &_
			" EmployeesCreditorsLKP.ConceptAmount, CreditorID, EmployeesCreditorsLKP.StartDate, EmployeesCreditorsLKP.EndDate," & _
			" CreditorTypeName, QttyName, Areas.AreaName, PaymentCenters.AreaName As Delegacion, Zones.ZoneCode, Zones.ZoneName," &_
			" EmployeesCreditorsLKP.ConceptMin, EmployeesCreditorsLKP.ConceptMax, CreditorsTypes.AppliesToID" & _
			" From Employees, EmployeesCreditorsLKP, CreditorsTypes, QttyValues, Areas, Areas As PaymentCenters, Zones" & _
			" Where (Employees.EmployeeID=EmployeesCreditorsLKP.EmployeeID)" & _
			" And (EmployeesCreditorsLKP.CreditorTypeID=CreditorsTypes.CreditorTypeID)" & _
			" And (CreditorsTypes.ConceptQttyID=QttyValues.QttyID)" & _
			" And (Areas.AreaID=EmployeesCreditorsLKP.PaymentCenterID)" & _
			" And (Areas.ParentID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID)"
			If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
				sQuery = sQuery & " And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sQuery = sQuery & " And (Employees.EmployeeID=0)"
				End If
			End If
			sQuery = sQuery & " And (EmployeesCreditorsLKP.Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")"
			If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then
				sQuery = sQuery & " And (EmployeesCreditorsLKP.StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
			End If
			sQuery = sQuery & " Order By Employees.EmployeeID, CreditorID, StartDate"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- " & Query & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If bForExport Then
				asColumnsTitles = Split("No. Empleado,Nombre Empleado,No. Acreedor,Nombre del acreedor,Fecha de inicio,Fecha de fin,Tipo de descuento,Cantidad de descuento,Unidades,Monto Mínimo,Monto Máximo,Centro de pago,Delegación,Municipio", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,",",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("No. Empleado,Nombre Empleado,No. Acreedor,Nombre del acreedor,Fecha de inicio,Fecha de fin,Tipo de descuento,Cantidad de descuento,Unidades,Monto Mínimo,Monto Máximo,Centro de pago,Delegación,Municipio,Acciones", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
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
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			asCellAlignments = Split(",,CENTER,,,,,CENTER,CENTER,CENTER,,", ",", -1, vbBinaryCompare)
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If bForExport Then
					sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
				Else
					sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
				If (Len(CStr(oRecordset.Fields("CreditorNumber").Value)) = 0) Or (CStr(oRecordset.Fields("CreditorNumber").Value) = "-1") Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditorNumber").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditorFullName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditorTypeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptMin").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptMax").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Delegacion").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) + "." + CStr(oRecordset.Fields("ZoneName").Value))
				lErrorNumber = GetConceptNamesFromAppliesToID(oADODBConnection, CStr(oRecordset.Fields("AppliesToID").Value), sConceptNames, sErrorDescription)
				iBeneficiaryID = CInt(oRecordset.Fields("CreditorID").Value)
				iStartDate = CLng(oRecordset.Fields("StartDate").Value)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(UCase(CStr(oRecordset.Fields("Reasons").Value)))
				If Not bForExport Then
					sRowContents = sRowContents & TABLE_SEPARATOR
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&BeneficiaryChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditorID=" & iBeneficiaryID & "&CreditorStartDate=" & iStartDate & "&ReasonID=" & lReasonID &""">"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Modificar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditorID=" & iBeneficiaryID & "&CreditorStartDate=" & iStartDate & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditorID=" & iBeneficiaryID & "&CreditorStartDate=" & iStartDate & "&ReasonID=" & lReasonID &""">"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					End If
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditorID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
					End If
				End If
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				bIsFirst = False
				iRecordCounter = iRecordCounter + 1
				If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Response.Write "</TABLE><BR /><BR />"
		Else
			If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El empleado seleccionado no tiene registros de beneficiarios de pensión alimenticia."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El empleado seleccionado no tiene registros en proceso de beneficiarios de pensión alimenticia."
				End If
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un número de empleado para buscar sus registros."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros de acreedores en proceso por ser aplicados"
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	DisplayEmployeesCreditorsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeChildrenTable(oRequest, oADODBConnection, sAction, lIDColumn, bUseLinks, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the employee's
'         children from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeChildrenTable"
	Dim sRequest
	Dim iRecordCounter
	Dim oRecordset
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
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesChildrenLKP.*, SchoolarshipName From EmployeesChildrenLKP, Schoolarships Where (EmployeesChildrenLKP.LevelID=Schoolarships.SchoolarshipID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By ChildBirthDate, ChildLastName, ChildLastName2, ChildName", "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = "Nombre,Fecha de nacimiento,Grado escolar de la beca"
				asCellWidths = "150,150,150"
				asCellAlignments = "CENTER,CENTER,CENTER"
				If Not(bForExport) Then
					asColumnsTitles = "Acciones," & asColumnsTitles
					asCellWidths = "100," & asCellWidths
					asCellAlignments = "CENTER," & asCellAlignments
				End If

				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ChildID").Value), oRequest("ChildID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CLng(oRecordset.Fields("ChildEndDate").Value) > 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ChildID"" ID=""ChildIDRd"" VALUE=""" & CStr(oRecordset.Fields("ChildID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ChildID"" ID=""ChildIDChk"" VALUE=""" & CStr(oRecordset.Fields("ChildID").Value) & """ />"
						Case Else
							If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
								sRowContents = sRowContents & "&nbsp;"
									If (CInt(Request.Cookies("SIAP_SubSectionID")) = 11) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 14) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 17) Or (CInt(Request.Cookies("SIAP_SubSectionID")) = 12) Then
										sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
									Else
										If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & oRequest("ReasonID") & "&ChildID=" & CStr(oRecordset.Fields("ChildID").Value) & "&Tab=1&Change=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
										End If
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (StrComp(sAction, "ChildrenSchoolarships", vbBinaryCompare) <> 0) Then
											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & oRequest("ReasonID") &  "&ChildID=" & CStr(oRecordset.Fields("ChildID").Value) & "&Tab=1&Delete=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
										End If
									End If
								sRowContents = sRowContents & "&nbsp;"
							End If
					End Select
					If not bForExport Then
						sRowContents = sRowContents & TABLE_SEPARATOR
					End If
					sRowContents = sRowContents & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ChildLastName").Value) & " " & CStr(oRecordset.Fields("ChildLastName2").Value) & " " & CStr(oRecordset.Fields("ChildName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CLng(oRecordset.Fields("ChildBirthDate").Value) > 0 Then 
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("ChildBirthDate").Value), -1, -1, -1)
							If CLng(oRecordset.Fields("ChildEndDate").Value) > 0 Then sRowContents = sRowContents & " - " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("ChildEndDate").Value), -1, -1, -1)
						Else
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("SchoolarshipName").Value)) & sBoldEnd & sFontEnd

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			Response.Write "<BR />- El empleado no tiene hijos registrados.<BR /><BR />"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeeChildrenTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingEmployeesBeneficiariesTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the beneficiaries of employees
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingEmployeesBeneficiariesTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
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
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion
	Dim iBeneficiaryID
	Dim iStartDate
	Dim bIsFirst
	Dim sConceptNames

	bIsFirst = True
	sErrorDescription = "No existen beneficiarios de pención alimenticia para el empledado."
	sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," &_
			" BeneficiaryNumber, BeneficiaryName + ' ' + BeneficiaryLastName + ' ' + BeneficiaryLastName2 As BeneficiaryFullName," &_
			" EmployeesBeneficiariesLKP.ConceptAmount, BeneficiaryID, EmployeesBeneficiariesLKP.StartDate, EmployeesBeneficiariesLKP.EndDate," & _
			" AlimonyTypeName, QttyName, Areas.AreaName, PaymentCenters.AreaName As Delegacion, Zones.ZoneCode, Zones.ZoneName," &_
			" EmployeesBeneficiariesLKP.ConceptMin, EmployeesBeneficiariesLKP.ConceptMax, AlimonyTypes.AppliesToID" & _
			" From Employees, EmployeesBeneficiariesLKP, AlimonyTypes, QttyValues, Areas, Areas As PaymentCenters, Zones" & _
			" Where (Employees.EmployeeID=EmployeesBeneficiariesLKP.EmployeeID)" & _
			" And (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID)" & _
			" And (AlimonyTypes.ConceptQttyID=QttyValues.QttyID)" & _
			" And (Areas.AreaID=EmployeesBeneficiariesLKP.PaymentCenterID)" & _
			" And (Areas.ParentID=PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID)"
			If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
				sQuery = sQuery & " And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					sQuery = sQuery & " And (Employees.EmployeeID=0)"
				End If
			End If
			sQuery = sQuery & " And (EmployeesBeneficiariesLKP.Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")"
			If False Then
				sQuery = sQuery & " And (EmployeesBeneficiariesLKP.StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
			End If
			sQuery = sQuery & " Order By Employees.EmployeeID, BeneficiaryID, StartDate"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If bForExport Then
				asColumnsTitles = Split("No. Empleado,Nombre Empleado,No. Beneficiario,Nombre del beneficiario,Fecha de inicio,Fecha de fin,Tipo de pensión,Descuento,Unidades,Monto Mínimo,Monto Máximo,Centro de pago,Delegación,Municipio", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,",",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("No. Empleado,Nombre Empleado,No. Beneficiario,Nombre del beneficiario,Fecha de inicio,Fecha de fin,Tipo de pensión,Descuento,Unidades,Monto Mínimo,Monto Máximo,Centro de pago,Delegación,Municipio,Acciones", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
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
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			asCellAlignments = Split(",,CENTER,,,,,CENTER,CENTER,CENTER,,", ",", -1, vbBinaryCompare)
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If bForExport Then
					sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
				Else
					sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
				If (Len(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) = 0) Or (CStr(oRecordset.Fields("BeneficiaryNumber").Value) = "-1") Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryFullName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptMin").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptMax").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Delegacion").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) + "." + CStr(oRecordset.Fields("ZoneName").Value))
				lErrorNumber = GetConceptNamesFromAppliesToID(oADODBConnection, CStr(oRecordset.Fields("AppliesToID").Value), sConceptNames, sErrorDescription)
				If False Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(sConceptNames))
				End If
				iBeneficiaryID = CInt(oRecordset.Fields("BeneficiaryID").Value)
				iStartDate = CLng(oRecordset.Fields("StartDate").Value)
				
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(UCase(CStr(oRecordset.Fields("Reasons").Value)))
				If Not bForExport Then
					sRowContents = sRowContents & TABLE_SEPARATOR
					If bIsFirst Then
						If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;"
								sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""10"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					Else
						If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&MoveBeneficiaryUp=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & iBeneficiaryID & "&ReasonID=" & lReasonID &""">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnCrclAddUp.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aumentar prioridad"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					End If
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&BeneficiaryChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & iBeneficiaryID & "&BeneficiaryStartDate=" & iStartDate & "&ReasonID=" & lReasonID &""">"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Modificar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & iBeneficiaryID & "&BeneficiaryStartDate=" & iStartDate & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & iBeneficiaryID & "&BeneficiaryStartDate=" & iStartDate & "&ReasonID=" & lReasonID &""">"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					End If
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("BeneficiaryID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
					End If
				End If
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				bIsFirst = False
				iRecordCounter = iRecordCounter + 1
				If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Response.Write "</TABLE><BR /><BR />"
		Else
			If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El empleado seleccionado no tiene registros de beneficiarios de pensión alimenticia."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El empleado seleccionado no tiene registros en proceso de beneficiarios de pensión alimenticia."
				End If
			Else
				If aEmployeeComponent(N_ACTIVE_EMPLOYEE) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un número de empleado para buscar sus registros."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros de beneficiarios de pensión alimenticia en proceso para ser aplicados"
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	DisplayPendingEmployeesBeneficiariesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingEmployeesTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, iActive, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the employees that are in process of movement
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingEmployeesTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
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
	Dim lErrorNumber
	Dim sCondition
	Dim sTables
	Dim sFields
	Dim sStatusEmployeesIDs
	Dim sDate
	Dim sQuery
	Dim bAux
	Dim iForPayrollIsActiveConstant
	Dim iRecordCounter

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sTables = ", Jobs"
		sCondition = "And (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	Else
		sTables = ""
		sCondition = ""
	End If
	sDate = CLng(Left(GetSerialNumberForDate(""), (Len("00000000"))))
	Select Case sAction
		Case "EmployeesMovements"
			If (lReasonID < -58) Then
				Select Case lReasonID
					Case -89
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 100
					Case EMPLOYEES_SAFE_SEPARATION
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120
					Case EMPLOYEES_ADD_SAFE_SEPARATION
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87
					Case EMPLOYEES_ANTIQUITIES
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 5
					Case EMPLOYEES_ADDITIONALSHIFT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
					Case EMPLOYEES_GLASSES
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 24
					Case EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 44						
					Case EMPLOYEES_FAMILY_DEATH
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 45
					Case EMPLOYEES_PROFESSIONAL_DEGREE
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 46
					Case EMPLOYEES_MONTHAWARD
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 50
					Case EMPLOYEES_SPORTS
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 165
					Case EMPLOYEES_SPORTS
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 69
					Case EMPLOYEES_CARLOAN
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 74
					Case EMPLOYEES_FOR_RISK
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
					Case -74
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
					Case EMPLOYEES_CONCEPT_08
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
					Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 22
					Case EMPLOYEES_LICENSES
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 104
					Case EMPLOYEES_CONCEPT_16
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 19
					Case EMPLOYEES_NON_EXCENT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 72
					Case EMPLOYEES_EXCENT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 73
					Case EMPLOYEES_MOTHERAWARD
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 26
					Case EMPLOYEES_HELP_COMISSION
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 63
					Case EMPLOYEES_SAFEDOWN
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 67
					Case EMPLOYEES_ANUAL_AWARD
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 32
					Case EMPLOYEES_BENEFICIARIES
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 70
					Case EMPLOYEES_BENEFICIARIES_DEBIT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
					Case EMPLOYEES_NIGHTSHIFTS
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 93
					Case EMPLOYEES_CONCEPT_C3
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 94
					Case EMPLOYEES_FONAC_ADJUSTMENT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 76
					Case EMPLOYEES_FONAC_CONCEPT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 77
					Case EMPLOYEES_CONCEPT_7S
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 146
                    Case EMPLOYEES_HONORARIUM_CONCEPT
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 13
				End Select
				If lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT Then
					sErrorDescription = "No existen créditos de terceros para el archivo."
					sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName, CreditID, CreditTypeShortName, CreditTypeName, PaymentAmount, BeneficiaryName + ' ' + BeneficiaryLastName + ' ' + BeneficiaryLastName2 As BeneficiaryFullName, Credits.StartDate, Credits.EndDate, Credits.Comments From Employees, Credits, CreditTypes, EmployeesBeneficiariesLKP" & sTables & " Where (Employees.EmployeeID=Credits.EmployeeID) And (Employees.EmployeeID=EmployeesBeneficiariesLKP.EmployeeID) And (((Credits.StartDate>=EmployeesBeneficiariesLKP.StartDate) And (Credits.StartDate<=EmployeesBeneficiariesLKP.EndDate)) Or ((Credits.EndDate>=EmployeesBeneficiariesLKP.StartDate) And (Credits.EndDate<=EmployeesBeneficiariesLKP.EndDate)) Or ((Credits.EndDate>=EmployeesBeneficiariesLKP.StartDate) And (Credits.StartDate<=EmployeesBeneficiariesLKP.EndDate))) And (Credits.BeneficiaryID=EmployeesBeneficiariesLKP.BeneficiaryID) And (Credits.CreditTypeID=CreditTypes.CreditTypeID) And (Credits.CreditTypeID=86) And (Credits.Active=" & iActive & ")"
					If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
						sQuery = sQuery & " And (Credits.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
					Else 
						sQuery = sQuery & " And (Credits.EmployeeID=0)"
					End If
					sQuery = sQuery & sContidion & " And (Credits.Active=" & iActive & ")"
					If iActive = 0 Then
						sQuery = sQuery & sContidion & " And (Credits.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
					End If
					sQuery = sQuery & sContidion & "Order By Employees.EmployeeID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				ElseIf (lReasonID = EMPLOYEES_ADDITIONALSHIFT) Or (lReasonID = EMPLOYEES_CONCEPT_08) Then
					sQuery = "Select EmployeesConceptsLKP.EmployeeID, EmployeeName, EmployeeLastName, " & _
								"EmployeeLastName2, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, " & _
								"ConceptShortName, ConceptName, ConceptAmount," & _
								"EmployeesConceptsLKP.EndUserID, EmployeesConceptsLKP.Comments, " & _
								"EmployeesConceptsLKP.ConceptID " & _
							"From EmployeesConceptsLKP, Concepts, Employees " & _
							"Where EmployeesConceptsLKP.EmployeeID = Employees.EmployeeID " & _
								"And EmployeesConceptsLKP.ConceptID = Concepts.ConceptID " & _
								"And Concepts.ConceptID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & _
								" And EmployeesConceptsLKP.Active = 0 " & _
								"And EmployeesConceptsLKP.EndDate > " & CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				ElseIf (lReasonID = EMPLOYEES_NIGHTSHIFTS) Then
					sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName, EmployeesConceptsLKP.ConceptID, ConceptShortName || '. ' || Concepts.ConceptName As Concept, EmployeesConceptsLKP.Active, ConceptAmount, ConceptQttyID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.StartUserID, EmployeesConceptsLKP.Comments, EmployeesConceptsLKP.RegistrationDate, QttyValues.QttyName, Users.Username + ' ' + Users.UserLastname As UserFullName From Employees, Concepts, EmployeesConceptsLKP, Users, QttyValues" & sTables & " Where (EmployeesConceptsLKP.StartUserID = Users.UserID) And (Employees.EmployeeID=EmployeesConceptsLKP.EmployeeID) And (Concepts.ConceptID=EmployeesConceptsLKP.ConceptID) And (QttyValues.QttyID = EmployeesConceptsLKP.ConceptQttyID) And (EmployeesConceptsLKP.ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & sCondition
					If CLng(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
						sQuery = sQuery & " And (EmployeesConceptsLKP.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
						'If iActive = 1 Then
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & iActive & ") Order By EmployeesConceptsLKP.EmployeeID, EmployeesConceptsLKP.StartDate Desc"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						'End If
					Else
						If iActive = 0 Then
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & iActive & ") Order By Employees.EmployeeID"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						Else
							sErrorDescription = "Para consultar los registros existentes debe buscar el número de empleado en la sección 1."
							lErrorNumber = -1
						End If
					End If
				Else
					sQuery = "Select Employees.EmployeeID, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesConceptsLKP.ConceptID, EmployeesConceptsLKP.Active, ConceptShortName, ConceptAmount, Concepts.ConceptName, ConceptQttyID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.StartUserID, EmployeesConceptsLKP.Comments, EmployeesConceptsLKP.RegistrationDate, QttyValues.QttyName, Users.Username, Users.UserLastname From Employees, Concepts, EmployeesConceptsLKP, Users, QttyValues" & sTables & " Where (EmployeesConceptsLKP.StartUserID = Users.UserID) And (Employees.EmployeeID=EmployeesConceptsLKP.EmployeeID) And (Concepts.ConceptID=EmployeesConceptsLKP.ConceptID) And (QttyValues.QttyID = EmployeesConceptsLKP.ConceptQttyID) And (EmployeesConceptsLKP.ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & sCondition
					If CLng(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
						sQuery = sQuery & " And (EmployeesConceptsLKP.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
						If iActive = 1 Then
							Select Case lReasonID
								Case EMPLOYEES_GLASSES, EMPLOYEES_PROFESSIONAL_DEGREE
									sQuery = sQuery & " And (EmployeesConceptsLKP.Active IN (1,2)) Order By Employees.EmployeeID"
								Case Else
									sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & iActive & ") Order By Employees.EmployeeID, EmployeesConceptsLKP.StartDate Desc"
							End Select
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						Else
							sQuery = sQuery & " And (EmployeesConceptsLKP.StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & iActive & ") Order By EmployeesConceptsLKP.EmployeeID, EmployeesConceptsLKP.StartDate Desc"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End If
					Else
						If iActive = 0 Then
							sQuery = sQuery & " And (EmployeesConceptsLKP.StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & iActive & ") Order By Employees.EmployeeID"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						Else
							sErrorDescription = "Para consultar los registros existentes debe buscar el número de empleado en la sección 1."
							lErrorNumber = -1
						End If
					End If
					aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
				End If
			Else
				Select Case lReasonID
					Case -58
						If iActive = 0 Then
							sCondition = sCondition & " And (EmployeesAdjustmentsLKP.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")"
						End If
						If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then aEmployeeComponent(N_ID_EMPLOYEE) = CLng(aEmployeeComponent(S_NUMBER_EMPLOYEE))
						If aEmployeeComponent(N_ID_EMPLOYEE) > 1 Then
							sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Reasons.ReasonName, EmployeesAdjustmentsLKP.ConceptID, Concepts.ConceptName, EmployeesAdjustmentsLKP.ConceptAmount, EmployeesAdjustmentsLKP.MissingDate, EmployeesAdjustmentsLKP.PayrollDate, EmployeesAdjustmentsLKP.BeneficiaryName, UserName, UserLastName, StartUserID From Concepts, Employees, EmployeesAdjustmentsLKP, Reasons, Users" & sTables & " Where (EmployeesAdjustmentsLKP.ConceptID=Concepts.ConceptID) And (EmployeesAdjustmentsLKP.EmployeeID=Employees.EmployeeID) And (EmployeesAdjustmentsLKP.UserID=Users.UserID) And (Reasons.ReasonID=" & lReasonID & ") And (EmployeesAdjustmentsLKP.PaymentDate=0) And (EmployeesAdjustmentsLKP.Active=" & iActive & ") And (EmployeesAdjustmentsLKP.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeesAdjustmentsLKP.MissingDate >= Concepts.StartDate) And (EmployeesAdjustmentsLKP.MissingDate <= Concepts.EndDate)" & sCondition
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						Else
							If iActive = 0 Then
								sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Reasons.ReasonName, EmployeesAdjustmentsLKP.ConceptID, Concepts.ConceptName, EmployeesAdjustmentsLKP.ConceptAmount, EmployeesAdjustmentsLKP.MissingDate, EmployeesAdjustmentsLKP.PayrollDate, EmployeesAdjustmentsLKP.BeneficiaryName, UserName, UserLastName From Concepts, Employees, EmployeesAdjustmentsLKP, Reasons, Users" & sTables & " Where (EmployeesAdjustmentsLKP.ConceptID=Concepts.ConceptID) And (EmployeesAdjustmentsLKP.EmployeeID=Employees.EmployeeID) And (EmployeesAdjustmentsLKP.UserID=Users.UserID) And (Reasons.ReasonID=" & lReasonID & ") And (EmployeesAdjustmentsLKP.PaymentDate=0) And (EmployeesAdjustmentsLKP.Active=" & iActive & ")" & sCondition
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							Else
								lErrorNumber = -1
								sErrorDescription = "Para consultar los registros existentes debe buscar el número de empleado en la sección 1."
							End If
						End If
						aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
					Case EMPLOYEES_FOR_RISK
					sQuery = "Select EmployeesConceptsLKP.EmployeeID, EmployeeName, EmployeeLastName, " & _
								"EmployeeLastName2, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, " & _
								"ConceptShortName, ConceptName, ConceptAmount," & _
								"EmployeesConceptsLKP.EndUserID, EmployeesConceptsLKP.Comments, " & _
								"EmployeesConceptsLKP.ConceptID " & _
							"From EmployeesConceptsLKP, Concepts, Employees " & _
							"Where EmployeesConceptsLKP.EmployeeID = Employees.EmployeeID " & _
								"And EmployeesConceptsLKP.ConceptID = Concepts.ConceptID " & _
								"And Concepts.ConceptID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & _
								" And EmployeesConceptsLKP.Active = 0 " & _
								"And EmployeesConceptsLKP.EndDate > " & CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					Case 54
						sQuery = "Select Jobs.JobID, Services.ServiceShortName, Services.ServiceName, Jobs.StartDate, Jobs.EndDate From Jobs, Services Where (Jobs.ServiceID = Services.ServiceID) And (Jobs.ModifyDate=" & sDate & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					Case 0
						sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.RFC, Employees.CURP, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeeTypes.EmployeeTypeName, Reasons.ReasonName From Employees, EmployeesHistoryList, EmployeeTypes, Reasons Where (Employees.StatusID=-2) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.ReasonID=0) Order by EmployeesHistoryList.EmployeeID"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					Case 12
						sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeeTypes.EmployeeTypeName, Reasons.ReasonName, StatusName From Employees, EmployeeTypes, EmployeesHistoryList, Reasons, StatusEmployees Where (Employees.StatusID = StatusEmployees.StatusID) And (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.StatusID In(-4, -3)) And (EmployeesHistoryList.bProcessed=2) And (EmployeesHistoryList.ReasonID=" & lReasonID & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					Case Else
						If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
							sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeeTypes.EmployeeTypeName, Reasons.ReasonName, StatusName From Employees, EmployeeTypes, EmployeesHistoryList, Reasons, StatusEmployees Where (Employees.StatusID = StatusEmployees.StatusID) And (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.StatusID Not In(0,1,26,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114,118,122,129,130)) And (EmployeesHistoryList.bProcessed=2) And (EmployeesHistoryList.ReasonID=" & lReasonID & ")"
						Else
							sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeeTypes.EmployeeTypeName, Reasons.ReasonName, StatusName From Employees, EmployeeTypes, EmployeesHistoryList, Jobs, Reasons, StatusEmployees Where (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))) And (Employees.StatusID = StatusEmployees.StatusID) And (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.StatusID Not In(0,1,26,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114,118,122,129,130)) And (EmployeesHistoryList.bProcessed=2) And (EmployeesHistoryList.ReasonID=" & lReasonID & ")"
						End If
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				End Select
			End If
	End Select
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"

	aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
	If lErrorNumber = 0 Then
		If iActive = 0 Then
			sErrorDescription = "No existen registros en proceso de aplicación."
		Else
			sErrorDescription = "No existen registros."
		End If
		If Not oRecordset.EOF Then
			If Not bForExport Then
				If lReasonID <= -58 Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			End If
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If (InStr(1, sAction, "EmployeesMovements", vbBinaryCompare) > 0) Then
				If (lReasonID < -58) Or (lReasonID = EMPLOYEES_FOR_RISK) Then 'Conceptos
					If bForExport Or (iActive = 1) Then
						If (lReasonID = EMPLOYEES_NIGHTSHIFTS) Then
							asColumnsTitles = Split("No. Emp.,Nombre,Concepto,Quincena de aplicación,Días festivos registrados,Usuario", ",", -1, vbBinaryCompare)
							'asCellWidths = Split("200,200,200,200,500,200",",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,CENTER,,,CENTER", ",", -1, vbBinaryCompare)
						ElseIf lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT Then
							asColumnsTitles = Split("No. Emp.,Nombre,Concepto,Importe,Beneficiario,Fecha inicio,Fecha fin,Observaciones", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,CENTER", ",", -1, vbBinaryCompare)
						ElseIf lReasonID = EMPLOYEES_ANTIQUITIES Then
							asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Fecha inicio,Fecha fin,Quincena de aplicación,Observaciones,Usuario", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,RIGHT,,,,", ",", -1, vbBinaryCompare)
						Else
							Select Case lReasonID
								Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS
									asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Quincena de aplicación,Observaciones,Usuario,Acciones", ",", -1, vbBinaryCompare)
									asCellWidths = Split(",,,,,,,,,",",", -1, vbBinaryCompare)
									asCellAlignments = Split(",,,,,,RIGHT,,,,", ",", -1, vbBinaryCompare)
								Case Else
									asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Fecha inicio,Fecha fin,Quincena de aplicación,Observaciones,Usuario", ",", -1, vbBinaryCompare)
									asCellWidths = Split(",,,,,,,,,,",",", -1, vbBinaryCompare)
									asCellAlignments = Split(",,,,,RIGHT,,,,", ",", -1, vbBinaryCompare)
							End Select
						End If
					Else
						If (lReasonID = EMPLOYEES_NIGHTSHIFTS) Then
							asColumnsTitles = Split("No. Emp.,Nombre,Concepto,Quincena de aplicación,Días festivos registrados,Usuario,Acciones", ",", -1, vbBinaryCompare)
							'asCellWidths = Split("200,200,200,200,500,200,200",",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,CENTER,,CENTER,CENTER,", ",", -1, vbBinaryCompare)
						ElseIf lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT Then
							asColumnsTitles = Split("No. Emp.,Nombre,Concepto,Importe,Beneficiario,Fecha inicio,Fecha fin,Observaciones, Acciones", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,CENTER", ",", -1, vbBinaryCompare)
						ElseIf lReasonID = EMPLOYEES_ANTIQUITIES Then
							asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Fecha inicio,Fecha fin,Quincena de aplicación,Observaciones,Usuario,Acciones", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,RIGHT,,,,,CENTER", ",", -1, vbBinaryCompare)
						ElseIf lReasonID = EMPLOYEES_FOR_RISK Or lReasonID = EMPLOYEES_CONCEPT_08 Or lReasonID = EMPLOYEES_ADDITIONALSHIFT Then
							asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Porcentaje,Fecha inicio,Fecha fin,Observaciones,Usuario,Acciones", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,,,",",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,RIGHT,,,,,", ",", -1, vbBinaryCompare)
						Else
							Select Case lReasonID
								Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS
									asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Quincena de aplicación,Observaciones,Usuario,Acciones", ",", -1, vbBinaryCompare)
									asCellWidths = Split(",,,,,,,,,",",", -1, vbBinaryCompare)
									asCellAlignments = Split(",,,,,RIGHT,,,,,CENTER", ",", -1, vbBinaryCompare)
								Case Else
									asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Clave de Concepto,Concepto,Importe,Fecha inicio,Fecha fin,Quincena de aplicación,Observaciones,Usuario,Acciones", ",", -1, vbBinaryCompare)
									asCellWidths = Split(",,,,,,,,,,",",", -1, vbBinaryCompare)
									asCellAlignments = Split(",,,,,,RIGHT,,,,,CENTER", ",", -1, vbBinaryCompare)
							End Select
						End If
					End If
				Else
					Select Case lReasonID
						Case -58
							If bForExport Or iActive = 1 Then
								asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Movimiento,Concepto,Importe,Quincena de aplicación,Fecha de omisión,Beneficiario,Usuario", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,,RIGHT,,,", ",", -1, vbBinaryCompare)
							Else
								asColumnsTitles = Split("No. Emp.,Apellido paterno,Apellido materno,Nombre,Movimiento,Concepto,Importe,Quincena de aplicación,Fecha de omisión,Beneficiario,Usuario,Acciones", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,,RIGHT,,,,CENTER", ",", -1, vbBinaryCompare)
							End If
						Case 0
							If bForExport Then
								asColumnsTitles = Split("No.Emp.,RFC,CURP,Apellido paterno,Apellido materno,Nombre,Movimiento,Tipo de tabulador", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,", ",", -1, vbBinaryCompare)
							Else
								asColumnsTitles = Split("No.Emp.,RFC,CURP,Apellido paterno,Apellido materno,Nombre,Movimiento,Tipo de tabulador,Acciones", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,,CENTER", ",", -1, vbBinaryCompare)
							End If
						Case 54
							asColumnsTitles = Split("No. plaza,Clave de servicio,Descripción del servicio", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
						Case Else
							If bForExport Or iActive = 1 Then
								asColumnsTitles = Split("No.Emp.,Apellido paterno,Apellido materno,Nombre,Movimiento,Tipo de tabulador,Estatus", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,,", ",", -1, vbBinaryCompare)
							Else
								asColumnsTitles = Split("No.Emp.,Apellido paterno,Apellido materno,Nombre,Movimiento,Tipo de tabulador,Estatus,Acciones", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,,", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)
							End If
					End Select
				End If
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
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				bAux = False
				Select Case sAction
					Case "EmployeesMovements"
						If (lReasonID = EMPLOYEES_FOR_RISK) Or (lReasonID = EMPLOYEES_ADDITIONALSHIFT) Or (lReasonID = EMPLOYEES_CONCEPT_08) Then
							sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & " "
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("ConceptAmount").Value)
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value))
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinido"
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value))
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EndUserID").Value))
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
								bAux = True
							End If
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ModifyConcept=1&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
								If bAux Then
									sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
								End If
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
							If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
							End If							
						ElseIf lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT Then
							If bForExport Then
								sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
							Else
								sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(UCase(CStr(oRecordset.Fields("CreditTypeName").Value)))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentAmount").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryFullName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
							End If
							If (Not IsNull(oRecordset.Fields("Comments").Value)) And (Len(oRecordset.Fields("Comments").Value) > 0) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(UCase(CStr(oRecordset.Fields("Comments").Value)))
							Else 
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
							End If
							If Not bForExport Then
								If (iActive = 0) Then
									iCreditID = CLng(oRecordset.Fields("CreditID").Value)
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(oRecordset.Fields("CreditID").Value) & "&ReasonID=" & lReasonID & """>"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(oRecordset.Fields("CreditID").Value) & "&ReasonID=" & lReasonID &""">"
												sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
									End If
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&CreditChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(oRecordset.Fields("CreditID").Value) & "&ReasonID=" & lReasonID &""">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Modificar registro"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
									If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And iActive=0 Then
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
									End If
								Else

								End If
							End If
						ElseIf (lReasonID < -58) Or (lReasonID = EMPLOYEES_FOR_RISK) Then 'Conceptos
							If lReasonID = EMPLOYEES_NIGHTSHIFTS Then
								If bForExport Then
									sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
								Else
									sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Concept").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value))
								Select Case lReasonID
									Case EMPLOYEES_NIGHTSHIFTS
										Dim sNightShiftsDates1
										Dim sNightShiftsDatesDesc1
										sNightShiftsDatesDesc1 = ""
										If Not IsEmpty(oRecordset.Fields("Comments").Value) Then
											sNightShiftsDates1 = Split(CStr(oRecordset.Fields("Comments").Value), ",", -1, vbBinaryCompare)
											For iIndex = 0 To UBound(sNightShiftsDates1)
												sNightShiftsDatesDesc1 = sNightShiftsDatesDesc1 & CStr(DisplayNumericDateFromSerialNumber(sNightShiftsDates1(iIndex))) & ","
											Next
											If InStr(1, Right(sNightShiftsDatesDesc1, Len(","), ",")) Then
												sNightShiftsDatesDesc1 = Left(sNightShiftsDatesDesc1, (Len(sNightShiftsDatesDesc1) - Len(",")))
											End If
										End If
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sNightShiftsDatesDesc1)
									Case Else
										If Len(CStr(oRecordset.Fields("Comments").Value)) = 0 Then
											sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Ninguna")
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
										End If
								End Select
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value))
							Else
								If bForExport Then
									sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
								Else
									sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & " "
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
								If (lReasonID = EMPLOYEES_FOR_RISK) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value))							
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
								End If
								If (lReasonID <> EMPLOYEES_FOR_RISK) Then
									If CLng(oRecordset.Fields("ConceptQttyID").Value) = 1 Then
										Select Case lReasonID
											Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
												sRowContents = sRowContents & TABLE_SEPARATOR & "NA"
											Case Else
												sRowContents = sRowContents & TABLE_SEPARATOR & "$ " & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
										End Select
									ElseIf CLng(oRecordset.Fields("ConceptQttyID").Value) = 2 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & " %"
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
									End If
								Else
										sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("RiskLevel").Value)*10, 2, True, False, True)
								End If
								If (lReasonID = EMPLOYEES_FOR_RISK) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value))
								Else
									Select Case lReasonID
										Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS
										Case Else
											sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
									End Select
								End If
								Select Case lReasonID
									Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS
									Case Else
										If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "A la fecha"
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
										End If
								End Select
								If (lReasonID <> EMPLOYEES_FOR_RISK) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & oRecordset.Fields("PayrollName").Value
								End If
								Select Case lReasonID
									Case EMPLOYEES_NIGHTSHIFTS
										Dim sNightShiftsDates
										Dim sNightShiftsDatesDesc
										If Not IsEmpty(oRecordset.Fields("Comments").Value) Then
											sNightShiftsDates = Split(CStr(oRecordset.Fields("Comments").Value), ",", -1, vbBinaryCompare)
											For iIndex = 0 To UBound(sNightShiftsDates)
												sNightShiftsDatesDesc = sNightShiftsDatesDesc & CStr(DisplayNumericDateFromSerialNumber(sNightShiftsDates(iIndex))) & ","
											Next
											If InStr(1, Right(sNightShiftsDatesDesc, Len(","), ",")) Then
												sNightShiftsDatesDesc = Left(sNightShiftsDatesDesc, (Len(sNightShiftsDatesDesc) - Len(",")))
											End If
										End If
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sNightShiftsDatesDesc)
									Case Else
										If Len(CStr(oRecordset.Fields("Comments").Value)) = 0 Then
											sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Ninguna")
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
										End If
								End Select
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("UserLastname").Value))
							End If
							If Not bForExport Then
								If (iActive = 0) Then
									If (lReasonID <> EMPLOYEES_FOR_RISK) Then
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											bAux = True
										End If
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&EmployeePayrollDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&ModifyConcept=1&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											If bAux Then
												sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
											Else
												sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
											End If
												sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
										End If
									Else
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&EmployeeDate=" & CStr(oRecordset.Fields("EmployeeDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											bAux = True
										End If
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&EmployeeDate=" & CStr(oRecordset.Fields("EmployeeDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											If bAux Then
												sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&EmployeeDate=" & CStr(oRecordset.Fields("EmployeeDate").Value) & "&ReasonID=" & lReasonID & """>"
											Else
												sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&EmployeeDate=" & CStr(oRecordset.Fields("EmployeeDate").Value) & "&ReasonID=" & lReasonID & """>"
											End If
												sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("StartDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
										End If
									End If
								Else
									'sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ModifyConcept=1&ReasonID=" & lReasonID & """>"
									'	sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									'sRowContents = sRowContents & "</A>&nbsp;"
									Select Case lReasonID
										Case EMPLOYEES_GLASSES, EMPLOYEES_PROFESSIONAL_DEGREE
											If CInt(oRecordset.Fields("Active").Value) = 1 Then
												sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&DeactiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											Else
												sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&ActiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										Case Else
									End Select
								End If
							End If
							lReasonID = CLng(oRecordset.Fields("ReasonID").Value)
							lStatusID = CLng(oRecordset.Fields("StatusID").Value)
						Else
							If lReasonID = 54 Then
								sRowContents = Right("000000" & CStr(oRecordset.Fields("JobID").Value), Len("000000"))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
							Else
								If CLng(oRecordset.Fields("EmployeeID").Value) >= 1000000 Then
									sRowContents = Right("0000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("0000000"))
								Else
									sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & " "
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value))
							End If
							Select Case lReasonID
								Case -58
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("PayrollDate").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("MissingDate").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("UserLastName").Value))
									If Not bForExport And iActive=0 Then
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&MissingDate=" & CStr(oRecordset.Fields("MissingDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar movimiento"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
											iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
										ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
											iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
										ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
											iForPayrollIsActiveConstant = 0
										End If
										If VerifyPayrollIsActive(oADODBConnection, CLng(oRecordset.Fields("PayrollDate").Value), iForPayrollIsActiveConstant, sErrorDescription) Then
											If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
												sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & "&SaveEmployeesMovements=1&Authorization=1&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&MissingDate=" & CStr(oRecordset.Fields("MissingDate").Value) & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar movimiento"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												'sRowContents = sRowContents & "&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("MissingDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" VALUE=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" />"
											End If
										Else
											sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
											sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
										End If
										If False Then
											If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
												sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&MissingDate=" & CStr(oRecordset.Fields("MissingDate").Value) & "&ReasonID=" & lReasonID & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Consultar el movimiento"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									End If
								Case 54
								Case 0
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value))
									If Not bForExport Then
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=58"">"
											'sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=58"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Reasignar número de empleado"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=58" & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Asignar plaza nuevo ingreso"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
									End If
								Case Else
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
									If Not bForExport Then
										If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Consultar el movimiento"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """/>"
									End If
							End Select
						End If
				End Select
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				iRecordCounter = iRecordCounter + 1
				If (Not bForExport) And (lReasonID <= -58) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Response.Write "</TABLE><BR /><BR />"
		Else
			If (lReasonID = EMPLOYEES_EXTRAHOURS) Or (lReasonID = EMPLOYEES_SUNDAYS) Or (lReasonID = EMPLOYEES_BENEFICIARIES_DEBIT) Then
				If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
					sErrorDescription = "El empleado no cuenta con registros del concepto"
				Else
					If iActive Then
						sErrorDescription = "Seleccione un número de empleado"
					Else
						sErrorDescription = "No existen registros en proceso"
					End If
				End If
			End If
			lErrorNumber = L_ERR_NO_RECORDS
		End If
	End If
	
	Set oRecordset = Nothing
	DisplayPendingEmployeesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the employees concepts that are in process of movement
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingEmployeesConceptsTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
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
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion
	Dim sConceptsIDs
	Dim iConceptsCount

	sErrorDescription = "No existen conceptos con estatus en proceso de autorización."
	Select Case lReasonID
		Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
			sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName , EmployeesAbsencesLKP.AbsenceID, AbsenceShortName, AbsenceHours, OcurredDate, AppliedDate" &_
					" From Employees, Absences, EmployeesAbsencesLKP" & _
					" Where " & _
					" (Employees.EmployeeID = EmployeesAbsencesLKP.EmployeeID)" & _
					" And (Absences.AbsenceID = EmployeesAbsencesLKP.AbsenceID)" & _
					" And (EmployeesAbsencesLKP.AbsenceID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")"
					If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
						sQuery = sQuery & " And (EmployeesAbsencesLKP.EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
					End IF
					sQuery = sQuery & " And (EmployeesAbsencesLKP.Active = " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")"
					sQuery = sQuery & " Order By Employees.EmployeeID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("N. de empleado,Nombre,Clave de Concepto, AppliedDate, Cantidad, Fecha, Estatus", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,",",", -1, vbBinaryCompare)
					Else
						If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							asColumnsTitles = Split("N. de empleado,Nombre,Clave de Concepto, Cantidad, Fecha de Ocurrencia, Fecha de Aplicacion, Acciones", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("N. de empleado,Nombre,Clave de Concepto, Cantidad, Fecha de Ocurrencia, Fecha de Aplicacion", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,", ",", -1, vbBinaryCompare)
						End If
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
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
					sFontBegin = ""
					sFontEnd = ""
					asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)
					Do While Not oRecordset.EOF
						If bForExport Then
							sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
						Else
							sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(DisplayNumericDateFromSerialNumber(oRecordset.Fields("OcurredDate").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(DisplayNumericDateFromSerialNumber(oRecordset.Fields("AppliedDate").Value))
						If (Not bForExport) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If (Not bForExport) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ConceptAmount=" & CStr(oRecordset.Fields("AbsenceHours").Value) & "&ReasonID=" & lReasonID &""">"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If (Not bForExport) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SundayChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					Response.Write "</TABLE><BR /><BR />"
				Else
					lErrorNumber = -1
					sErrorDescription = "No existen empleados en proceso de movimiento."
				End If
			End If
		Case Else
			If lReasonID = CANCEL_EMPLOYEES_CONCEPTS Then
				sConceptsIDs = EMPLOYEES_CONCEPTS
			ElseIf lReasonID = CANCEL_EMPLOYEES_C04 Then
				If (CInt(Request.Cookies("SIAP_SubSectionID")) = 1) Then
					sConceptsIDs = "4,7,8"
				Else
					sConceptsIDs = "72,73,100"
				End If
			ElseIf lReasonID = CANCEL_EMPLOYEES_SSI Then
				sConceptsIDs = BENEFIT_CONCEPTS_FOR_SSI
			Else
				sConceptsIDs = BENEFIT_CONCEPTS_FOR_PERSONAL
			End If
			sQuery = "Select Employees.EmployeeID, EmployeeName, EmployeeLastName, EmployeesConceptsLKP.ConceptID, ConceptShortName, ConceptName, ConceptAmount, ConceptQttyID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.RegistrationDate, EmployeesConceptsLKP.Active " &_
					" From Employees, Concepts, EmployeesConceptsLKP" & _
					" Where " & _
					" (Employees.EmployeeID = EmployeesConceptsLKP.EmployeeID)" & _
					" And (Concepts.ConceptID = EmployeesConceptsLKP.ConceptID)" & _
					" And (EmployeesConceptsLKP.EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
'						" And (((EmployeesConceptsLKP.StartDate >= " &  Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EmployeesConceptsLKP.EndDate <= " &  Left(GetSerialNumberForDate(""), Len("00000000")) & "))" & _
'						" Or ((EmployeesConceptsLKP.EndDate >= " &  Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EmployeesConceptsLKP.EndDate <=30000000))" & _
'						" Or ((EmployeesConceptsLKP.EndDate >= " &  Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EmployeesConceptsLKP.StartDate <=30000000)))"
					If (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Then
						If CInt(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) = 2 Then
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & ") And (EmployeesConceptsLKP.EndDate<>0)"
						Else
							sQuery = sQuery & " And (EmployeesConceptsLKP.Active=" & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & ")"
						End If
					Else
						sQuery = sQuery & " And EmployeesConceptsLKP.Active=1"
					End If
					sQuery = sQuery & " And Concepts.ConceptID IN (" & sConceptsIDs & ")" & _
					" Order By EmployeesConceptsLKP.ConceptID, EmployeesConceptsLKP.StartDate"
			Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					iConceptsCount = 0
					Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno, Clave de Concepto, Descripción, Cantidad, Fecha inicial, Fecha final, Quincena de aplicación, Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,",",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno, Clave de Concepto, Descripción, Cantidad, Fecha inicial, Fecha final, Quincena de aplicación, Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
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
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
					sFontBegin = ""
					sFontEnd = ""
					asCellAlignments = Split(",,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					Do While Not oRecordset.EOF
						iConceptsCount = iConceptsCount + 1
						If bForExport Then
							sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
						Else
							sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "A la fecha"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value))
						If CInt(oRecordset.Fields("Active").Value) = 2 Then
							'sRowContents = sRowContents & TABLE_SEPARATOR & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
							Select Case (CLng(oRecordset.Fields("StartDate").Value) = CLng(oRecordset.Fields("EndDate").Value))
								Case True
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										If (CLng(oRecordset.Fields("StartDate").Value) > CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) Or (CLng(oRecordset.Fields("StartDate").Value) < CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) And CLng(oRecordset.Fields("RegistrationDate").Value)) > CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&ActiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Re-Activar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
										End If
									End If
							End Select
						Else
							Select Case (CLng(oRecordset.Fields("StartDate").Value) = CLng(oRecordset.Fields("EndDate").Value))
								Case True
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ConceptEndDate=" & CStr(oRecordset.Fields("EndDate").Value) & "&EmployeePayrollDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&ReasonID=" & lReasonID & """>"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								Case False
									If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&EmployeePayrollDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&ModifyConcept=1&ReasonID=" & lReasonID & """>"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Dar de baja"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&EmployeePayrollDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&ModifyConcept=1&Cancel=1&ReasonID=" & lReasonID & """>"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
							End Select
						End If
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					Select Case lReasonID
						Case CANCEL_EMPLOYEES_CONCEPTS
							Response.Write "</TABLE><BR />"
							If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
								If CInt(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) = 1 Then
									Call DisplayInstructionsMessage("Número de registros", "El empleado cuenta con:&nbsp;" & iConceptsCount & " registros de prestaciones.")
								Else
									Call DisplayInstructionsMessage("Número de registros", "El empleado cuenta con:&nbsp;" & iConceptsCount & " prestaciones canceladas.")
								End If
							Else
								Call DisplayInstructionsMessage("Número de registros", "Intrduzca un No. Empleado.")
							End If
						Case Else
							Response.Write "</TABLE><BR />"
					End Select
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
						If (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) And (CInt(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) = 2) Then
							sErrorDescription = "El empleado no cuenta con registros de conceptos cancelados"
						Else
							sErrorDescription = "El empleado no cuenta con registros de conceptos activos"
						End If
					Else
						sErrorDescription = "Seleccione un número de empleado"
					End If
				End If
			End If
	End Select
	Set oRecordset = Nothing
	DisplayPendingEmployeesConceptsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesFeaturesTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the employees concepts that are in process of movement
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesFeaturesTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
	Dim oRecordset1
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
	Dim lErrorNumber
	Dim lErrorNumber1
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion
	Dim sConceptsIDs
	Dim iConceptsCount

	sConceptsIDs = EMPLOYEES_CONCEPTS

	sErrorDescription = "No existen conceptos con estatus en proceso de autorización."

	sQuery = "Select Employees.EmployeeID, EmployeeName, EmployeeLastName, EmployeesAbsencesLKP.AbsenceID, AbsenceShortName, AbsenceName, AbsenceHours, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, AppliedDate" &_
			" From Employees, Absences, EmployeesAbsencesLKP" & _
			" Where " & _
			" Employees.EmployeeID = EmployeesAbsencesLKP.EmployeeID" & _
			" And Absences.AbsenceID = EmployeesAbsencesLKP.AbsenceID" & _
			" And EmployeesAbsencesLKP.AbsenceID IN (201,202)" & _
			" And EmployeesAbsencesLKP.OcurredDate >= " & Left(GetSerialNumberForDate(""), Len("00000000"))
			If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
				sQuery = sQuery & " and EmployeesAbsencesLKP.EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE)
			End IF
			sQuery = sQuery & " Order By Employees.EmployeeID"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

	sQuery = "Select Employees.EmployeeID, EmployeeName, EmployeeLastName, EmployeesConceptsLKP.ConceptID, ConceptShortName, ConceptName, ConceptAmount, ConceptQttyID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.RegistrationDate, EmployeesConceptsLKP.Active " &_
			" From Employees, Concepts, EmployeesConceptsLKP" & _
			" Where " & _
			" (Employees.EmployeeID = EmployeesConceptsLKP.EmployeeID)" & _
			" And (Concepts.ConceptID = EmployeesConceptsLKP.ConceptID)" & _
			" And (EmployeesConceptsLKP.EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
			" And ((EmployeesConceptsLKP.EndDate<=30000000) And (EmployeesConceptsLKP.EndDate>=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ")" & _
			" Or (EmployeesConceptsLKP.StartDate > " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")" & _
			" Or ((EmployeesConceptsLKP.StartDate < " & Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EmployeesConceptsLKP.RegistrationDate> " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")))"
	sQuery = sQuery & " And Concepts.ConceptID IN (" & sConceptsIDs & ")" & _
			" Order By EmployeesConceptsLKP.ConceptID, EmployeesConceptsLKP.StartDate"
	lErrorNumber1 = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)

	lErrorNumber1 = 0
	If (lErrorNumber = 0) And (lErrorNumber1 = 0) Then
		Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
			asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno,Clave de Concepto, AppliedDate, Cantidad, Fecha, Estatus", ",", -1, vbBinaryCompare)
			asCellWidths = Split(",,,,,,,",",", -1, vbBinaryCompare)
		Else
			asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno,Clave de Concepto, Cantidad, Fecha de Ocurrencia, Fecha de Aplicacion, Acciones", ",", -1, vbBinaryCompare)
			asCellWidths = Split(",,,,,,,", ",", -1, vbBinaryCompare)
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
		sBoldBegin = "<B>"
		sBoldEnd = "</B>"
		sFontBegin = ""
		sFontEnd = ""
		asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)

		If (Not oRecordset.EOF) And (oRecordset1.EOF) Then
			If Not oRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
					asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno,Clave de Concepto, AppliedDate, Cantidad, Fecha, Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split(",,,,,,,",",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Número de empleado,Nombre,Apellido paterno,Clave de Concepto, Cantidad, Fecha de Ocurrencia, Fecha de Aplicacion, Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split(",,,,,,,", ",", -1, vbBinaryCompare)
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
				sBoldBegin = "<B>"
				sBoldEnd = "</B>"
				sFontBegin = ""
				sFontEnd = ""
				asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If bForExport Then
						sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
					Else
						sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CStr(oRecordset.Fields("OcurredDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CStr(oRecordset.Fields("AppliedDate").Value))
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ConceptAmount=" & CStr(oRecordset.Fields("AbsenceHours").Value) & "&ReasonID=" & lReasonID &""">"
							sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SundayChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
					sFontEnd = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontBegin = "</FONT>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			End If
			Do While Not oRecordset.EOF
				iConceptsCount = iConceptsCount + 1
				If bForExport Then
					sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
				Else
					sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value))
				If CInt(oRecordset.Fields("Active").Value) = 2 Then
					Select Case (CLng(oRecordset.Fields("StartDate").Value) = CLng(oRecordset.Fields("EndDate").Value))
						Case True
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								If (CLng(oRecordset.Fields("StartDate").Value) > CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) Or (CLng(oRecordset.Fields("StartDate").Value) < CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) And CLng(oRecordset.Fields("RegistrationDate").Value)) > CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&ActiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Re-Activar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
								End If
							End If
					End Select
				Else
					Select Case (CLng(oRecordset.Fields("StartDate").Value) = CLng(oRecordset.Fields("EndDate").Value))
						Case True
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						Case False
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ModifyConcept=1&ReasonID=" & lReasonID & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
					End Select
				End If
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Response.Write "</TABLE><BR />"
		Else
			Response.Write "</TABLE><BR />"
			lErrorNumber = -1
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				sErrorDescription = "Introduzca el número de empleado para consultar sus prestaciones vigentes."
			Else
				sErrorDescription = "El empleado no tiene prestaciones vigentes."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "Error al verificar si el empleado cuenta con prestaciones vigentes."
	End If

	Set oRecordset = Nothing
	DisplayEmployeesFeaturesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, iActive, bForExport, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the employees credits of employees
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingEmployeesCreditsTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
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
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion
	Dim iCreditID
	Dim iRecordType
	Dim iRecordCounter

	If lReasonID = EMPLOYEES_THIRD_PROCESS Then
		sErrorDescription = "No existen créditos de terceros para el archivo."
		sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," &_
				" CreditID, Credits.CreditTypeID, CreditTypeShortName, CreditTypeName, StartAmount, PaymentAmount, PaymentsNumber, UploadedFileName, UploadedRecordType, Comments, QttyName" & _
				" From Employees, Credits, CreditTypes, QttyValues" & _
				" Where (Employees.EmployeeID=Credits.EmployeeID)" & _
				" And (Credits.CreditTypeID=CreditTypes.CreditTypeID)" & _
				" And (Credits.QttyID = QttyValues.QttyID)" & _
				" And (Credits.Active=0)"
				If Len(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE)) > 0 Then
					sQuery = sQuery & " And (UploadedFileName='" & aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) & "')"
				Else
					sQuery = sQuery & " And (UploadedFileName='9999')"
				End If
				sQuery = sQuery & " Order By Employees.EmployeeID, CreditTypeID"
	Else
		sErrorDescription = "No existen créditos de terceros con estatus en proceso de autorización."
		sQuery = "Select Employees.EmployeeID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," &_
				" CreditID, Credits.CreditTypeID, CreditTypeShortName, CreditTypeName, StartAmount, PaymentAmount, PaymentsNumber, Credits.StartDate, Credits.EndDate, Comments, QttyName" & _
				" From Employees, Credits, CreditTypes, QttyValues" & _
				" Where (Employees.EmployeeID = Credits.EmployeeID)" & _
				" And (Credits.CreditTypeID = CreditTypes.CreditTypeID)" & _
				" And (Credits.QttyID = QttyValues.QttyID)" & _
				" And ((UploadedFileName='') Or (UploadedFileName Is Null))"
				If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
					sQuery = sQuery & " And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
				Else
					sQuery = sQuery & " And (Employees.EmployeeID=0)"
				End If
				sQuery = sQuery & " And (Credits.Active=" & iActive & ")"
				sQuery = sQuery & " Order By Employees.EmployeeID, CreditTypeID"
	End If
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If False Then
			'If lReasonID = EMPLOYEES_THIRD_PROCESS Then
				If bForExport Or (iActive = 1) Then
					If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Observaciones,Archivo de Tercero,Tipo de  registro", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,,",",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Observaciones,Archivo de Tercero,Tipo de  registro", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
					End If
				Else
					If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Observaciones,Archivo de Tercero,Tipo de  registro,Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,",",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Observaciones,Archivo de Tercero,Tipo de  registro,Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
					End If
				End If
			Else
				If bForExport Then
				'If bForExport Or (iActive = 1) Then
					If False Then
					'If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Fecha Inicio,Fecha Término,Observaciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,",",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Fecha Inicio,Fecha Termino,Observaciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					End If
				Else
					If False Then
					'If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Fecha Inicio,Fecha Término,Observaciones,Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,",",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("N. de empleado,Nombre,Clave Concepto,Descripción,Cuota fija, Unidad, Número de pagos,Importe,Fecha Inicio,Fecha Término,Observaciones,Acciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					End If
				End If
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
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditTypeShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditTypeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentAmount").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentsNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StartAmount").Value))
				If lReasonID <> EMPLOYEES_THIRD_PROCESS Then
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "A la fecha"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
					End If
				End If
				If Len(CStr(oRecordset.Fields("Comments").Value)) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Ninguna")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
				End If
				If lReasonID = EMPLOYEES_THIRD_PROCESS Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UploadedFileName").Value))
					Select Case CInt(oRecordset.Fields("UploadedRecordType").Value)
						Case 1
							iRecordType = "Alta"
						Case 3
							iRecordType = "Baja"
						Case 2
							iRecordType = "Cambio"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(iRecordType))
				End If
				If Not bForExport Then
					If iActive = 0 Then
						iCreditID = CLng(oRecordset.Fields("CreditID").Value)
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(iCreditID) & "&ReasonID=" & lReasonID & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(iCreditID) & "&ReasonID=" & lReasonID &""">"
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
						If lReasonID <> EMPLOYEES_THIRD_PROCESS Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&CreditChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(iCreditID) & "&ReasonID=" & lReasonID &""">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Modificar registro"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
						'If lReasonID = EMPLOYEES_THIRD_PROCESS Then
						'	If B_DELETE And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And iActive=0  Then
						'		sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
						'	End If
						'End If
					Else
						iCreditID = CLng(oRecordset.Fields("CreditID").Value)
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS And (CLng(oRecordset.Fields("EndDate").Value) > CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) Then
							'sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&BeneficiaryChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&BeneficiaryID=" & iBeneficiaryID & "&BeneficiaryStartDate=" & iStartDate & "&ReasonID=" & lReasonID &""">"
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&CreditChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&CreditID=" & CStr(iCreditID) & "&ReasonID=" & lReasonID & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					End If
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If				
				End If
				oRecordset.MoveNext
				If ((Err.number <> 0) Or (lErrorNumber <> 0)) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
			Loop
			Response.Write "</TABLE><BR /><BR />"
		Else
			If lReasonID = EMPLOYEES_THIRD_PROCESS Then
				If Len(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE)) > 0 Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El archivo seleccionado no tiene registros de terceros."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un archivo para buscar sus registros."
				End If			
			Else
				If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El empleado seleccionado no tiene registros de terceros."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un número de empleado para buscar sus registros."
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	DisplayPendingEmployeesCreditsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayConcepts040708HistoryList(oRequest, oADODBConnection, bForExport, bFull, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the history list for the employee from
'         the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConcepts040708HistoryList"
	Dim oRecordset
	Dim sCondition
	Dim sBoldBegin
	Dim sBoldEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim lConcept

	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "EmployeeDate", False, sCondition)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And (Areas.AreaID Like (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	If lReasonID = EMPLOYEES_FOR_RISK Then lConcept = 4
	If lReasonID = EMPLOYEES_CONCEPT_08 Then lConcept = 8
	If lReasonID = EMPLOYEES_ADDITIONALSHIFT Then lConcept = 7
	
	sQuery = "Select EmployeeID, EmployeesConceptsLKP.ConceptID, ConceptName, ConceptAmount," & _
			"EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate " & _
			"From EmployeesConceptsLKP, Concepts " & _
			"Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID)" & _
			" And (EmployeesConceptsLKP.ConceptID=" & lConcept & ")" & _
			" And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
			" And (EmployeesConceptsLKP.Active=1)" & _
			" Order By ConceptID, EndDate Desc"
	
	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery , "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If bFull Then
				Response.Write "<TABLE WIDTH=""3000"" BORDER="""
			Else
				Response.Write "<TABLE BORDER="""
			End If
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = "Concepto,Monto,Fecha inicio,Fecha fin"
				asCellWidths = "150,100,150,150"
				asCellAlignments = ",RIGHT,CENTER,CENTER"
				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				aEmployeeComponent(N_HISTORY_LIST_RECORTS) = 0
				Do While Not oRecordset.EOF
					aEmployeeComponent(N_HISTORY_LIST_RECORTS) = aEmployeeComponent(N_HISTORY_LIST_RECORTS) + 1
					sBoldBegin = ""
					sBoldEnd = ""
					If (CLng(oRequest("StartDate").Item) = CLng(oRecordset.Fields("StartDate").Value)) Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber(oRecordset.Fields("ConceptAmount").Value) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sBoldEnd
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Indefinida" & sBoldEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)) & sBoldEnd
					Err.Clear
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = "-1"
			sErrorDescription = "No existen registros que cumplan con el criterio de la búsqueda"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayConcepts040708HistoryList = lErrorNumber
	Err.Clear
End Function
%>