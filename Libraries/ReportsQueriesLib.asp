<!-- #include file="ReportsQueries1000Lib.asp" -->
<!-- #include file="ReportsQueries1000bLib.asp" -->
<!-- #include file="ReportsQueries1100Lib.asp" -->
<!-- #include file="ReportsQueries1100bLib.asp" -->
<!-- #include file="ReportsQueries1200Lib.asp" -->
<!-- #include file="ReportsQueries1300Lib.asp" -->
<!-- #include file="ReportsQueries1400Lib.asp" -->
<!-- #include file="ReportsQueries1400bLib.asp" -->
<!-- #include file="ReportsQueries1400cLib.asp" -->
<!-- #include file="ReportsQueries1500Lib.asp" -->
<!-- #include file="ReportsQueries1600Lib.asp" -->
<%
Function DisplayLogCount(oRequest, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the users log count
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayLogCount"
	Dim sCondition
	Dim lCounter
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetStartAndEndDatesFromURL("StartLog", "EndLog", "LogDate", True, sCondition)

	sErrorDescription = "No se pudieron obtener las entradas al sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(SystemLogs.UserID) As LogsCount, UserName, UserLastName From Users, SystemLogs Where (Users.UserID=SystemLogs.UserID) " & sCondition & " Group By UserName, UserLastName Order By UserName, UserLastName", "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""550"" BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Usuario,Entradas", ",", -1, vbBinaryCompare)
				asCellWidths = Split("300,250", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("LogsCount").Value), 0, False, False, True)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					lCounter = lCounter + CLng(oRecordset.Fields("LogsCount").Value)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "<B>Número de entradas</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(lCounter, 0, False, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen entradas al sistema en el rango de fechas especificado."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayLogCount = lErrorNumber
	Err.Clear
End Function

Function DisplayAreasCount(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To count the areas using the filter information
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreasCount"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim iIndex
	Dim oRecordset
	Dim iCounter
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If (InStr(1, " " & sTableNames & ",", " Areas,", vbBinaryCompare) = 0) Then sTableNames = "Areas, " & sTableNames
	sJoinCondition = Replace(sJoinCondition, "Employees.", "Areas.")
	sJoinCondition = Replace(sJoinCondition, "Jobs.", "Areas.")
	If Len(sSortFields) = 0 Then sSortFields = "Areas.AreaID"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Areas.AreaID " & sFieldNames & " From " & sTableNames & " Where (AreaID>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, 0, "Centro de trabajo"), ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				iCounter = 0
				lErrorNumber = BuildRowData(oRecordset, 1, -1, sCurrentRecords)
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sTempRecords)
					If StrComp(sTempRecords, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sRowContents = sCurrentRecords
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						iCounter = 0
						sCurrentRecords = sTempRecords
					End If
					iCounter = iCounter + 1
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sRowContents = Replace(Replace(sCurrentRecords, LIST_SEPARATOR, TABLE_SEPARATOR), "&#38;nbsp;", "&nbsp;")
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen centros de trabajo registrados en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAreasCount = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesCount(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To count the employees using the filter information
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesCount"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim iIndex
	Dim oRecordset
	Dim iCounter
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If ((InStr(1, sCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".AreaID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".AreaID", vbBinaryCompare) > 0)) And (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & ", Jobs"
	If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) > 0) And ((InStr(1, sJoinCondition, "(Employees.JobID=Jobs.JobID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Jobs.JobID=Employees.JobID)", vbBinaryCompare) = 0)) Then sJoinCondition = sJoinCondition & " And (Employees.JobID=Jobs.JobID)"
	If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) = 0) Then sTableNames = "Employees, " & sTableNames
	If Len(sSortFields) = 0 Then sSortFields = "Employees.EmployeeID"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Employees.EmployeeID " & sFieldNames & " From " & sTableNames & " Where (EmployeeID>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, 0, "Centro de trabajo"), ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				iCounter = 0
				lErrorNumber = BuildRowData(oRecordset, 1, -1, sCurrentRecords)
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sTempRecords)
					If StrComp(sTempRecords, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sRowContents = sCurrentRecords
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						iCounter = 0
						sCurrentRecords = sTempRecords
					End If
					iCounter = iCounter + 1
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sRowContents = Replace(Replace(sCurrentRecords, LIST_SEPARATOR, TABLE_SEPARATOR), "&#38;nbsp;", "&nbsp;")
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen centros de trabajo registrados en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesCount = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesList(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees data stored in the database.
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesList"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim oRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	lErrorNumber = GetConditionFromURL(oRequest, sCondition, -1, -1)
	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If ((InStr(1, sCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".AreaID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".AreaID", vbBinaryCompare) > 0)) And (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & ", Jobs"
	If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) > 0) And ((InStr(1, sJoinCondition, "(Employees.JobID=Jobs.JobID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Jobs.JobID=Employees.JobID)", vbBinaryCompare) = 0)) Then sJoinCondition = sJoinCondition & " And (Employees.JobID=Jobs.JobID)"
	If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) = 0) Then sTableNames = "Employees, " & sTableNames
	If Len(sSortFields) = 0 Then sSortFields = "Employees.EmployeeID"
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then
			sTableNames = sTableNames & ", Jobs"
		End If
		sCondition = sCondition & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	If (InStr(1, " " & sJoinCondition & ",", "Areas", vbBinaryCompare) <> 0) Then
		If (InStr(1, " " & sTableNames & ",", "Areas", vbBinaryCompare) = 0) Then 
			sTableNames = sTableNames & ", Areas"
			sCondition = sCondition & " And (Areas.ZoneID=Jobs.ZoneID)"
		End If
	End If

	sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Employees.EmployeeID " & sFieldNames & " From " & sTableNames & " Where (Employees.EmployeeID>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, "", ""), ",")
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sTempRecords)
					If StrComp(sTempRecords, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sRowContents = Replace(sTempRecords, LIST_SEPARATOR, TABLE_SEPARATOR)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						sCurrentRecords = sTempRecords
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayEmployeesList = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsCount(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To count the jobs using the filter information
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsCount"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim iIndex
	Dim oRecordset
	Dim iCounter
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If ((InStr(1, sCondition, ".CompanyID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".CompanyID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".EmployeeTypeID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".EmployeeTypeID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".PositionTypeID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".PositionTypeID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".GroupGradeLevelID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".GroupGradeLevelID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".JourneyID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".JourneyID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".LevelID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".LevelID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "Employees.StatusID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, "Employees.StatusID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".PaymentCenterID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".PaymentCenterID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".GenderID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".GenderID", vbBinaryCompare) > 0)) And (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & ", Employees"
	If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = "Jobs, " & sTableNames
	If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) > 0) And ((InStr(1, sJoinCondition, "(Employees.JobID=Jobs.JobID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Jobs.JobID=Employees.JobID)", vbBinaryCompare) = 0)) Then sJoinCondition = sJoinCondition & " And (Employees.JobID=Jobs.JobID)"
	If Len(sSortFields) = 0 Then sSortFields = "Jobs.JobID"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Jobs.JobID " & sFieldNames & " From " & sTableNames & " Where (Jobs.JobID>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, 0, "Centro de trabajo"), ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				iCounter = 0
				lErrorNumber = BuildRowData(oRecordset, 1, -1, sCurrentRecords)
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sTempRecords)
					If StrComp(sTempRecords, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sRowContents = sCurrentRecords
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						iCounter = 0
						sCurrentRecords = sTempRecords
					End If
					iCounter = iCounter + 1
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sRowContents = Replace(Replace(sCurrentRecords, LIST_SEPARATOR, TABLE_SEPARATOR), "&#38;nbsp;", "&nbsp;")
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(iCounter, 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen centros de trabajo registrados en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayJobsCount = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsList(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the jobs data stored in the database.
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsList"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim oRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iRecordCounter
	Dim iStartPage

	lErrorNumber = GetConditionFromURL(oRequest, sCondition, -1, -1)
	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If CLng(oRequest("ReportID").Item) <> 712 Then
		If CLng(oRequest("ReportID").Item) = 711 Then
			If InStr(1,sCondition,"JobsHistoryList") > 0 Then
				sTableNames = sTableNames & ", JobsHistoryList"
				sJoinCondition = sJoinCondition & " And (Jobs.JobID = JobsHistoryList.JobID)"
			End If
		End If	
		If Len(oRequest("MovedEmployees").Item) > 0 Then
			If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) = 0) Then sTableNames = "Employees, " & sTableNames
		End If
		'If ((InStr(1, sCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".ZoneID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".PositionID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".JobTypeID", vbBinaryCompare) > 0) Or (InStr(1, sCondition, ".AreaID", vbBinaryCompare) > 0) Or (InStr(1, sJoinCondition, ".AreaID", vbBinaryCompare) > 0)) And (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & ", Jobs"
		If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) > 0) And ((InStr(1, sJoinCondition, "(Employees.JobID=Jobs.JobID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Jobs.JobID=Employees.JobID)", vbBinaryCompare) = 0)) Then sJoinCondition = sJoinCondition & " And (Employees.JobID=Jobs.JobID)"
		If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = "Jobs, " & sTableNames
		If Len(sSortFields) = 0 Then sSortFields = "Jobs.JobID"

		If Len(oRequest("MovedEmployees").Item) > 0 Then
			sCondition = sCondition & " And (Employees.StatusID In (1,31,32,33))"
		ElseIf Len(oRequest("JobsOwners").Item) > 0 Then
			sCondition = sCondition & " And (Jobs.OwnerID>0)"
		End If

		sErrorDescription = "No se pudo obtener la información de las plazas registradas en el sistema."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Jobs.JobID " & sFieldNames & " From " & sTableNames & " Where (Jobs.JobID>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, OwnerID, CompanyName, Areas.AreaName, PaymentCenters.AreaName, PositionName, Positions.LevelID, Positions.classificationID, Positions.GroupGradeLevelID, Positions.IntegrationID, JobTypeName, OccupationTypeName, ServiceName, JourneyName, Jobs.WorkingHours From Jobs, Companies, Areas, Areas As PaymentCenters, Positions, JobTypes, OccupationTypes, Services, Journeys where (Jobs.CompanyID = Companies.CompanyID) And (Jobs.AreaID = Areas.AreaID) And (Jobs.PaymentCenterID = PaymentCenters.AreaID) And (Jobs.PositionID = Positions.PositionID) And  (Jobs.JobTypeID = JobTypes.JobTypeID) And (Jobs.OccupationTypeID = OccupationTypes.OccupationTypeID) And (Jobs.ServiceID = Services.ServiceID) And (Jobs.JourneyID = Journeys.JourneyID) And (Jobs.JobID > -1) " & Replace(sCondition, "XXX", "Jobs.Modify") & " Order By Jobs.ModifyDate Desc, jobID Desc" , "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If CLng(oRequest("ReportID").Item) <> 712 Then
					asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, "", ""), ",")
				Else
					asColumnsTitles = Split("Plaza,Titular,Empresa,Adscripción, Centro de Pago, Puesto, Nivel, Clasificación, Grupo grado nivel, Integración, Tipo de plaza, Tipo de ocupación, Servicio, Turno, Jornada", ",", -1, vbBinaryCompare)
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

				asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
				iStartPage = 1
				If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
				If Not bForExport Then Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
				iRecordCounter = 0
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sTempRecords)
					If StrComp(sTempRecords, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sRowContents = Replace(sTempRecords, LIST_SEPARATOR, TABLE_SEPARATOR)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						sCurrentRecords = sTempRecords
					End If
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_CATALOG) Then Exit Do
					If Err.number <> 0 Then Exit Do
					
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayJobsList = lErrorNumber
	Err.Clear
End Function
%>