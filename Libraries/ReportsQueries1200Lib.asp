<%
Function BuildReport1200(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal con conceptos. Reporte basado en la hoja 001221 
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1200"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim oCompaniesRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sHeaderContents
	Dim asConceptTitle
	Dim bEmpty
	Dim lTotalEmployees
	Dim dTotalAmount
	Dim iCompany

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If (InStr(1, sCondition, "Concepts.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Concepts.", "Percepciones.")
	End If

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	lTotalEmployees = 0
	dTotalAmount = 0
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		If lErrorNumber = 0 Then
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()		
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1024.htm"), sErrorDescription)
			If Len(sHeaderContents) > 0 Then
				asConceptTitle = Split(aReportTitle(L_CONCEPT_ID_FLAGS), ";")
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NAME />", asConceptTitle(1))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NUMBER />", asConceptTitle(0))
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
			End If
			lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
			If Not (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyName From Companies Where ParentID >=0 And EndDate=30000000 Order By CompanyShortName", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCompaniesRecordset)
				If lErrorNumber = 0 Then
					If Not oCompaniesRecordset.EOF Then
						Do While Not oCompaniesRecordset.EOF
							iCompany = CInt(oCompaniesRecordset.Fields("CompanyID").Value)
							sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								lTotalEmployees = 0
								dTotalAmount = 0
								If Not oRecordset.EOF Then
									bEmpty = False
									sRowContents = "<BR /><B>EMPRESA:" & Cstr(oRecordset.Fields("CompanyName").Value) &  "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										Else
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										End If
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
										sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lTotalEmployees = 0
									dTotalAmount = 0
									bEmpty = False
									sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										Else
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										End If
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
										sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " OPERATIVOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							End If
							oCompaniesRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oCompaniesRecordset.Close
					End If
				End If
			Else
				sCondition = Replace(sCondition, "Companies.", "Employees.")
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						bEmpty = False
						lTotalEmployees = 0
						dTotalAmount = 0
						sRowContents = "<BR /><B>Empresa: " & aReportTitle(L_COMPANY_FLAGS) & "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeTypeID, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						bEmpty = False
						sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lTotalEmployees = 0
						dTotalAmount = 0
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				End If
			End If
			If Not bEmpty Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			oRecordset.Close
		End If
	End If
	Set oRecordset = Nothing
	BuildReport1200 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1201(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal con conceptos. Reporte basado en la hoja 001221 
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1201"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim oCompaniesRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sHeaderContents
	Dim asConceptTitle
	Dim bEmpty
	Dim lTotalEmployees
	Dim dTotalAmount
	Dim iCompany

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If (InStr(1, sCondition, "Concepts.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Concepts.", "Percepciones.")
	End If

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	lTotalEmployees = 0
	dTotalAmount = 0
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		If lErrorNumber = 0 Then
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()		
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1024.htm"), sErrorDescription)
			If Len(sHeaderContents) > 0 Then
				asConceptTitle = Split(aReportTitle(L_CONCEPT_ID_FLAGS), ";")
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NAME />", asConceptTitle(1))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NUMBER />", asConceptTitle(0))
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
			End If
			lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
			If Not (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyName From Companies Where (ParentID>=0) And (EndDate=30000000) Order By CompanyShortName", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCompaniesRecordset)
				If lErrorNumber = 0 Then
					If Not oCompaniesRecordset.EOF Then
						Do While Not oCompaniesRecordset.EOF
							iCompany = CInt(oCompaniesRecordset.Fields("CompanyID").Value)
							sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, EmployeeTypes.EmployeeTypeShortName From Employees, EmployeeTypes, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID=Jobs.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) and (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								lTotalEmployees = 0
								dTotalAmount = 0
								If Not oRecordset.EOF Then
									bEmpty = False
									sRowContents = "<BR /><B>EMPRESA:" & Cstr(oRecordset.Fields("CompanyName").Value) &  "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "<TD>TIPO</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										Else
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										End If
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lTotalEmployees = 0
									dTotalAmount = 0
									bEmpty = False
									sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										Else
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										End If
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
										sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " OPERATIVOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							End If
							oCompaniesRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oCompaniesRecordset.Close
					End If
				End If
			Else
				sCondition = Replace(sCondition, "Companies.", "Employees.")
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						bEmpty = False
						lTotalEmployees = 0
						dTotalAmount = 0
						sRowContents = "<BR /><B>Empresa: " & aReportTitle(L_COMPANY_FLAGS) & "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeTypeID, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						bEmpty = False
						sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lTotalEmployees = 0
						dTotalAmount = 0
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				End If
			End If
			If Not bEmpty Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			oRecordset.Close
		End If
	End If
	Set oRecordset = Nothing
	BuildReport1201 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1201b(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal con conceptos. Reporte basado en la hoja 001221 
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1201b"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim oCompaniesRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sHeaderContents
	Dim asConceptTitle
	Dim bEmpty
	Dim lTotalEmployees
	Dim dTotalAmount
	Dim iCompany

	Dim asStateNames
	Dim asPath
	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim bFirst
	Dim lTotalEmployeesForArea
	Dim lTotalForArea
	Dim sTempCurrent
	Dim sCompanyName

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If (InStr(1, sCondition, "Concepts.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Concepts.", "Percepciones.")
	End If

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	lTotalEmployees = 0
	dTotalAmount = 0
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		If lErrorNumber = 0 Then
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			'sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()		
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1024.htm"), sErrorDescription)
			If Len(sHeaderContents) > 0 Then
				asConceptTitle = Split(aReportTitle(L_CONCEPT_ID_FLAGS), ";")
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NAME />", asConceptTitle(1))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NUMBER />", asConceptTitle(0))
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
			End If
			lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)

			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				asStateNames = ""
				Do While Not oRecordset.EOF
					asStateNames = asStateNames & LIST_SEPARATOR & SizeText(CStr(CleanStringForHTML(oRecordset.Fields("ZoneName").Value)), " ", 19, 1)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				asStateNames = Split(asStateNames, LIST_SEPARATOR)
			End If

			If Not (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyName From Companies Where (ParentID>=0) And (EndDate=30000000) Order By CompanyShortName", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCompaniesRecordset)
				If lErrorNumber = 0 Then
					If Not oCompaniesRecordset.EOF Then
						Do While Not oCompaniesRecordset.EOF
							iCompany = CInt(oCompaniesRecordset.Fields("CompanyID").Value)
							sCompanyName = Cstr(oCompaniesRecordset.Fields("CompanyName").Value)
							lCurrentPaymentCenterID = -2
							bFirst = False
							sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, EmployeeTypes.EmployeeTypeShortName From Employees, EmployeeTypes, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID=Jobs.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) and (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, EmployeeTypes.EmployeeTypeShortName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName From Employees, EmployeeTypes, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID=Jobs.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) and (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID)" & sCondition & " And Employees.CompanyID=" & iCompany & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, EmployeeTypes.EmployeeTypeShortName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName From Employees, EmployeeTypes, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID=Jobs.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) and (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID)" & sCondition & " And Employees.CompanyID=" & iCompany & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sRowContents = "<BR /><B>EMPRESA:" & sCompanyName &  "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
										If (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
											If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
												' Aquí van a ir los totales cuando es cambio de Centro de trabajo
												sRowContents = "</TABLE>"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												lTotalEmployeesForArea=0
												lTotalForArea=0
											End If
											' Se pinta la nueva sección
											If Len(asPath(2)) > 0 Then
												sRowContents = "<BR /><B>DELEGACION ESTATAL: " & CStr(asStateNames(CInt(asPath(2)))) & "</B><BR /><BR />"
											Else
												sRowContents = "<BR /><B>DELEGACION ESTATAL: (-1) NINGUNA</B><BR /><BR />"
											End If
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
											sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
											sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
											sRowContents = sRowContents & "<TD>No.EMP.</TD>"
											sRowContents = sRowContents & "<TD>NOMBRE</TD>"
											sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
											sRowContents = sRowContents & "<TD>RFC</TD>"
											sRowContents = sRowContents & "<TD>PUESTO</TD>"
											sRowContents = sRowContents & "<TD>N/SN</TD>"
											sRowContents = sRowContents & "<TD>MONTO</TD>"
											sRowContents = sRowContents & "<TD>TIPO</TD>"
											sRowContents = sRowContents & "</TR></FONT>"
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										End If
										lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
										sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
										sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										Else
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										End If
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "</FONT></TR>"
										lTotalEmployees = lTotalEmployees + 1
										lTotalEmployeesForArea = lTotalEmployeesForArea + 1
										lTotalForArea = lTotalForArea + CDbl(oRecordset.Fields("ConceptAmount").Value)
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										bFirst = True
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><BR /><B>TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lTotalEmployees = 0
									dTotalAmount = 0
									oRecordset.Close
								End If
							End If
							lTotalEmployeesForArea=0
							lTotalForArea=0
							lCurrentPaymentCenterID = -2
							bFirst = False
							sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sRowContents = "<BR /><B>EMPRESA:" & sCompanyName &  "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
										If (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
											If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
												' Aquí van a ir los totales cuando es cambio de Centro de trabajo
												sRowContents = "</TABLE>"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
												lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
												lTotalEmployeesForArea=0
												lTotalForArea=0
											End If
											' Se pinta la nueva sección
											If Len(asPath(2)) > 0 Then
												sRowContents = "<BR /><B>DELEGACION ESTATAL: " & CStr(asStateNames(CInt(asPath(2)))) & "</B><BR /><BR />"
											Else
												sRowContents = "<BR /><B>DELEGACION ESTATAL: (-1) NINGUNA</B><BR /><BR />"
											End If
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
											sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
											sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
											sRowContents = sRowContents & "<TD>No.EMP.</TD>"
											sRowContents = sRowContents & "<TD>NOMBRE</TD>"
											sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
											sRowContents = sRowContents & "<TD>RFC</TD>"
											sRowContents = sRowContents & "<TD>PUESTO</TD>"
											sRowContents = sRowContents & "<TD>N/SN</TD>"
											sRowContents = sRowContents & "<TD>MONTO</TD>"
											sRowContents = sRowContents & "</TR></FONT>"
											lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										End If
										lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
										sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
										sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										Else
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										End If
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "</FONT></TR>"
										lTotalEmployees = lTotalEmployees + 1
										lTotalEmployeesForArea = lTotalEmployeesForArea + 1
										lTotalForArea = lTotalForArea + CDbl(oRecordset.Fields("ConceptAmount").Value)
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										bFirst = True
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><BR /><B>TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lTotalEmployees = 0
									dTotalAmount = 0
									oRecordset.Close
								End If
							End If
							oCompaniesRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oCompaniesRecordset.Close
					End If
				End If
			Else
				sCondition = Replace(sCondition, "Companies.", "Employees.")
				lCurrentPaymentCenterID = -2
				bFirst = False
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sRowContents = "<BR /><B>EMPRESA:" & sCompanyName &  "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Do While Not oRecordset.EOF
							asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
							If (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
								If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
									' Aquí van a ir los totales cuando es cambio de Centro de trabajo
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lTotalEmployeesForArea=0
									lTotalForArea=0
								End If
								' Se pinta la nueva sección
								If Len(asPath(2)) > 0 Then
									sRowContents = "<BR /><B>DELEGACION ESTATAL: " & CStr(asStateNames(CInt(asPath(2)))) & "</B><BR /><BR />"
								Else
									sRowContents = "<BR /><B>DELEGACION ESTATAL: (-1) NINGUNA</B><BR /><BR />"
								End If
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>No.EMP.</TD>"
								sRowContents = sRowContents & "<TD>NOMBRE</TD>"
								sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
								sRowContents = sRowContents & "<TD>RFC</TD>"
								sRowContents = sRowContents & "<TD>PUESTO</TD>"
								sRowContents = sRowContents & "<TD>N/SN</TD>"
								sRowContents = sRowContents & "<TD>MONTO</TD>"
								sRowContents = sRowContents & "<TD>TIPO</TD>"
								sRowContents = sRowContents & "</TR></FONT>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							End If
							lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
							sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
							sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
							Else
								sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
							End If
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lTotalEmployees = lTotalEmployees + 1
							lTotalEmployeesForArea = lTotalEmployeesForArea + 1
							lTotalForArea = lTotalForArea + CDbl(oRecordset.Fields("ConceptAmount").Value)
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							bFirst = True
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><BR /><B>TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees & "</B>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "</B><BR /><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lTotalEmployees = 0
						dTotalAmount = 0
						oRecordset.Close
					End If
				End If

				lCurrentPaymentCenterID = -2
				bFirst = False
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeNumber, Employees.PaymentCenterID, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) " & sCondition & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sRowContents = "<BR /><B>EMPRESA:" & sCompanyName &  "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Do While Not oRecordset.EOF
							asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
							If (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
								If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
									' Aquí van a ir los totales cuando es cambio de Centro de trabajo
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lTotalEmployeesForArea=0
									lTotalForArea=0
								End If
								' Se pinta la nueva sección
								If Len(asPath(2)) > 0 Then
									sRowContents = "<BR /><B>DELEGACION ESTATAL: " & CStr(asStateNames(CInt(asPath(2)))) & "</B><BR /><BR />"
								Else
									sRowContents = "<BR /><B>DELEGACION ESTATAL: (-1) NINGUNA</B><BR /><BR />"
								End If
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>No.EMP.</TD>"
								sRowContents = sRowContents & "<TD>NOMBRE</TD>"
								sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
								sRowContents = sRowContents & "<TD>RFC</TD>"
								sRowContents = sRowContents & "<TD>PUESTO</TD>"
								sRowContents = sRowContents & "<TD>N/SN</TD>"
								sRowContents = sRowContents & "<TD>MONTO</TD>"
								sRowContents = sRowContents & "</TR></FONT>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							End If
							lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
							sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
							sTempCurrent = CStr(oRecordset.Fields("EmployeeID").Value)
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
							Else
								sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
							End If
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lTotalEmployees = lTotalEmployees + 1
							lTotalEmployeesForArea = lTotalEmployeesForArea + 1
							lTotalForArea = lTotalForArea + CDbl(oRecordset.Fields("ConceptAmount").Value)
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							bFirst = True
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>NUMERO DE FUNCIONARIOS: " & lTotalEmployeesForArea & "</B>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>MONTO DE FUNCIONARIOS $ " & FormatNumber(lTotalForArea, 2, True, False, True) & "</B><BR /><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><BR /><B>TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees & "</B>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "</B><BR /><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lTotalEmployees = 0
						dTotalAmount = 0
						oRecordset.Close
					End If
				End If
			End If
			If Not bEmpty Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			oRecordset.Close
		End If
	End If
	Set oRecordset = Nothing
	BuildReport1201b = lErrorNumber
	Err.Clear
End Function

Function BuildReport1202(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de créditos a los empleados
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1202"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim sCondition
	Dim sCondition2

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asCreditsNames
	Dim asPath
	Dim iCount
	Dim aiCreditsTotals
	Dim aiCreditsGrandTotals
	Dim iIndex
	Dim sCreditShortName
	Dim bFirst
	Dim lTotal
	Dim iMin
	Dim iMax

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "CreditStart", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "CreditEnd", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("CreditStart", "CreditEnd", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "Credits.Start") & ") Or (" & Replace(sCondition2, "XXX", "Credits.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "Credits.End", 1, 1, vbBinaryCompare), "XXX", "Credits.Start", 1, 1, vbBinaryCompare) & "))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		asStateNames = ""
		Do While Not oRecordset.EOF
			asStateNames = asStateNames & LIST_SEPARATOR & SizeText(CStr(CleanStringForHTML(oRecordset.Fields("ZoneName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asStateNames = Split(asStateNames, LIST_SEPARATOR)
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(CreditTypeID) As Max From CreditTypes", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value)
		End If
	End If
	For iMin = 0 To iMax
		asCreditsNames = asCreditsNames & LIST_SEPARATOR & ""
		aiCreditsTotals = aiCreditsTotals & LIST_SEPARATOR & "0"
		aiCreditsGrandTotals = aiCreditsGrandTotals & LIST_SEPARATOR & "0"
	Next
	asCreditsNames = Split(asCreditsNames, LIST_SEPARATOR)
	aiCreditsTotals = Split(aiCreditsTotals, LIST_SEPARATOR)
	aiCreditsGrandTotals = Split(aiCreditsGrandTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiCreditsTotals(iIndex) = 0
		aiCreditsGrandTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CreditTypeID, CreditTypeShortName From CreditTypes Order By CreditTypeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asCreditsNames(CInt(oRecordset.Fields("CreditTypeID").Value)) = SizeText(CStr(CleanStringForHTML(oRecordset.Fields("CreditTypeShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de créditos indicados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.PaymentCenterID, Employees.EmployeeNumber, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, CreditID, StartAmount, PaymentsNumber, PaymentAmount, DebtAmount, Credits.CreditTypeID, CreditTypeShortName, CreditTypeName, Credits.StartDate, Credits.EndDate, Credits.Active, CreditTypeName, PaymentAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName, Users.UserName + ' ' + Users.UserLastName As UserFullName From Employees, Credits, CreditTypes, Jobs, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, Companies, Users Where (Employees.EmployeeID=Credits.EmployeeID) And (Credits.CreditTypeID=CreditTypes.CreditTypeID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Credits.UserID=Users.UserID)" & sCondition & sCondition2, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeID, Employees.PaymentCenterID, Employees.EmployeeNumber, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, CreditID, StartAmount, PaymentsNumber, PaymentAmount, DebtAmount, Credits.CreditTypeID, CreditTypeShortName, CreditTypeName, Credits.StartDate, Credits.EndDate, Credits.Active, CreditTypeName, PaymentAmount, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName, Users.UserName + ' ' + Users.UserLastName As UserFullName From Employees, Credits, CreditTypes, Jobs, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, Companies, Users Where (Employees.EmployeeID=Credits.EmployeeID) And (Credits.CreditTypeID=CreditTypes.CreditTypeID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Credits.UserID=Users.UserID)" & sCondition & sCondition2 & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1108.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
					lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				End If
				iCount = 0
				lCurrentPaymentCenterID = -2
				bFirst = False
				Do While Not oRecordset.EOF
					iCount = iCount + 1
					asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
					If (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
						If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
							sRowContents = "</TABLE>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
							sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>CLAVE DEL CREDITO</TD>"
							sRowContents = sRowContents & "<TD>TOTAL</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							For iIndex = 0 To UBound(aiCreditsTotals)
								lTotal = CInt(aiCreditsTotals(iIndex))
								If lTotal > 0 Then
									sCreditShortName = Trim(asCreditsNames(CInt(iIndex)))
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & sCreditShortName & "</TD>"
										sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
									sRowContents = sRowContents & "</FONT></TR>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							Next
							For iIndex = 0 To UBound(aiCreditsTotals)
								aiCreditsTotals(iIndex) = 0
							Next
							sRowContents = "</TABLE>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
						End If
						If Len(asPath(2)) > 0 Then
							sRowContents = "<BR /><B>DELEGACION ESTATAL: " & CStr(asStateNames(CInt(asPath(2)))) & "</B><BR /><BR />"
						Else
							sRowContents = "<BR /><B>DELEGACION ESTATAL: (-1) NINGUNA</B><BR /><BR />"
						End If
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. Emp.</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave de crédito</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad inicial</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. de Pagos</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cuota fija</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad restante</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Estatus</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Usuario que registro</FONT></TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("EmployeeNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("EmployeeFullName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("CreditTypeShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("CreditTypeName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("StartAmount").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("PaymentsNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("PaymentAmount").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("DebtAmount").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("A la fecha") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("Activo") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("En proceso") & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML((oRecordset.Fields("UserFullName").Value)) & "</FONT></TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					aiCreditsTotals(CInt(oRecordset.Fields("CreditTypeID").Value)) = aiCreditsTotals(CInt(oRecordset.Fields("CreditTypeID").Value)) + 1
					aiCreditsGrandTotals(CInt(oRecordset.Fields("CreditTypeID").Value)) = aiCreditsGrandTotals(CInt(oRecordset.Fields("CreditTypeID").Value)) + 1
					bFirst = True
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					sRowContents = "<BR /><B>TOTALES POR CENTRO DE TRABAJO: " & sCurrentPaymentCenterName & "</B><BR />"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
					sRowContents = sRowContents & "<TD>CLAVE DEL CREDITO</TD>"
					sRowContents = sRowContents & "<TD>TOTAL</TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					For iIndex = 0 To UBound(aiCreditsTotals)
						lTotal = CInt(aiCreditsTotals(iIndex))
						If lTotal > 0 Then
							sCreditShortName = Trim(asCreditsNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sCreditShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
					For iIndex = 0 To UBound(aiCreditsTotals)
						aiCreditsTotals(iIndex) = 0
					Next
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
				End If
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>TOTALES DEL REPORTE</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
				sRowContents = sRowContents & "<TD>CLAVE DEL CREDITO</TD>"
				sRowContents = sRowContents & "<TD>TOTAL</TD>"
				sRowContents = sRowContents & "</FONT></TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				For iIndex = 0 To UBound(aiCreditsGrandTotals)
					lTotal = CInt(aiCreditsGrandTotals(iIndex))
					If lTotal > 0 Then
						sCreditShortName = Trim(asCreditsNames(iIndex))
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & sCreditShortName & "</TD>"
							sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				Next
				For iIndex = 0 To UBound(aiCreditsTotals)
					aiCreditsTotals(iIndex) = 0
				Next
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				oRecordset.Close
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1202 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1203(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de créditos a los empleados
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1203"
	Dim sHeaderContents
	Dim sHeaderContentsForEmployee
	Dim oRecordset
	Dim oRecordset1
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim sQuery
	Dim sCondition
	Dim sCondition2
	Dim sEmployeeName
	Dim sEmployeesNumbers
	Dim asEmployeesNumbers
	Dim sFolderPath
	Dim sZipFile
	Dim sActiveEmployeesStatus
	Dim sCancelEmployeesStatus
	Dim lStartDate
	Dim lEndDate
	Dim sStartDate
	Dim sEndDate
	Dim sReasonName
	Dim sEndReasonName
	Dim iHistoryCount
	Dim lHistoryStartDate
	Dim lHistoryEndDate
	Dim lLastHistoryStartDate
	Dim dConceptAntiquityAmount
	Dim dAnotherConceptAmount
	Dim sComments
	Dim sDocumentTypeCondition
	Dim sDocumentTypeConditionForCV
	Dim sDocumentSuffix
	Dim asDocumentSuffix
	Dim iDocIndex

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asCreditsNames
	Dim asPath
	Dim iCount
	Dim aiCreditsTotals
	Dim aiCreditsGrandTotals
	Dim iIndex
	Dim sCreditShortName
	Dim bFirst
	Dim lTotal
	Dim iMin
	Dim iMax

	sDocumentSuffix = "_01,_02,_03,_04,_05,_06"
	asDocumentSuffix = Split(sDocumentSuffix, ",", -1, vbBinaryCompare)
	sActiveEmployeesStatus = "-1,0,1,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,27,28,29,31,32,33,35,36,37,39,40,41,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,119,120,121,123,124,125,126,127,128,130,131,132,133,134,135,136,137,138,139,140,141,142,143,145,146,147,149,150,151,152,153,154,155,156,157,158"
	sCancelEmployeesStatus = "2,3,4,26,30,34,38,42,118,122,129,144,148,159"

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "DocumentStart", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "DocumentEnd", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("DocumentStart", "DocumentEnd", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	'If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesDocs.DocumentStart") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesDocs.DocumentEnd") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesDocs.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesDocs.DocumentStart", 1, 1, vbBinaryCompare) & "))"
	If Len(sCondition2) > 0 Then sCondition2 = " And " & Replace(sCondition2, "XXX", "EmployeesDocs.Document")
	If Len(oRequest("DocumentTypeID").Item) > 0 Then
		If CInt(oRequest("DocumentTypeID").Item) = 1 Then
			Dim lDocumentTypeStartDate, lDocumentTypeEndDate
			lDocumentTypeEndDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
			lDocumentTypeStartDate = AddYearsToSerialDate(Left(GetSerialNumberForDate(""), Len("00000000")), -1)
			sDocumentTypeCondition = " And (((EHL.EmployeeDate>=" & lDocumentTypeStartDate & ") And (EHL.EmployeeDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((EHL.EndDate>=" & lDocumentTypeStartDate & ") And (EHL.EndDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((EHL.EndDate>=" & lDocumentTypeStartDate & ") And (EHL.EmployeeDate<=" & lDocumentTypeEndDate & ")))"
			sDocumentTypeConditionForCV = " And (((ConceptsValues.StartDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.StartDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((ConceptsValues.EndDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.EndDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((ConceptsValues.EndDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.StartDate<=" & lDocumentTypeEndDate & ")))"
		End If
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesDocs.EmployeeID From EmployeesDocs, Employees Where (EmployeesDocs.EmployeeID=Employees.EmployeeID)" & sCondition & sCondition2 & " Order By EmployeesDocs.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sEmployeesNumbers = ""
		If Not oRecordset.EOF Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesServicesSheet", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesServicesSheetAmounts", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1203.rtf"), sErrorDescription)
			Do While Not oRecordset.EOF
				sComments = CStr(oRecordset.Fields("Comments").Value)
				sEmployeesNumbers = sEmployeesNumbers & CStr(oRecordset.Fields("EmployeeID").Value) & NUMERIC_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			sEmployeesNumbers = Left(sEmployeesNumbers, (Len(sEmployeesNumbers) - Len(",")))
			oRecordset.Close
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen empleados registrados para la entrega de hoja única de servicios para los criterios indicados."
		End If
	End If
	If lErrorNumber = 0 Then
		oStartDate = Now()
		sErrorDescription = "No se pudieron obtener los registros del empleado."
		sCondition2 = " And (Employees.EmployeeID IN (" & sEmployeesNumbers & "))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheet Select Distinct EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID, 0 As CashOfficer, 0 As Concept01Amount, 0 As Concept06Amount, 0 As AnotherConceptAmount From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, StatusEmployees, Reasons, Companies, EmployeeTypes, Employees Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.PositionID=Positions.PositionID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.EmployeeID=Employees.EmployeeID) And (EHL.EmployeeDate<=EHL.EndDate)" & sCondition2 & sDocumentTypeCondition & "Group by EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesServicesSheet Select Distinct EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID, 0 As CashOfficer, 0 As Concept01Amount, 0 As Concept06Amount, 0 As AnotherConceptAmount From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, StatusEmployees, Reasons, Companies, EmployeeTypes, Employees Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.PositionID=Positions.PositionID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.EmployeeID=Employees.EmployeeID) And (EHL.EmployeeDate<=EHL.EndDate)" & sCondition2 & sDocumentTypeCondition & "Group by EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc" & " -->" & vbNewLine
		If lErrorNumber = 0 Then
'-->> Ojo
			sDate = GetSerialNumberForDate("")
			sFolderPath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
			lErrorNumber = CreateFolder(Server.MapPath(sFolderPath), sErrorDescription)
			sZipFile = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			'Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sZipFile) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheet Where (StatusID IN(" & sActiveEmployeesStatus & ")) Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			iHistoryCount = 0
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						iHistoryCount = iHistoryCount + 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 1 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 1 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV & " -->" & vbNewLine
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
'Nuevo
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheet Where (StatusID IN(" & sCancelEmployeesStatus & ")) Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			'iHistoryCount = 0
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						iHistoryCount = iHistoryCount + 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 0 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 0 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV & " -->" & vbNewLine
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheetAmounts Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						lStartDate = oRecordset.Fields("StartDate").Value
						lEndDate = oRecordset.Fields("EndDate").Value
						aEmployeeComponent(N_ID_EMPLOYEE) = oRecordset.Fields("EmployeeID").Value
						lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, lStartDate, sEmployeeAntiquity, lAntiquityYears, lAntiquityMonths, lAntiquityDays, sErrorDescription)
						lAntiquityYears = lAntiquityYears + lAntiquityMonths/12
						'Call GetConceptAntiquityAmount(lAntiquity, dConceptAntiquityAmount)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from ConceptsValues Where (ConceptID=6) And (Antiquity2ID<=" & lAntiquityYears & ") And (EmployeeTypeID=" & oRecordset.Fields("EmployeeTypeID").Value & ") And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) Order By Antiquity2ID Desc", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						If lErrorNumber = 0 Then
							If Not oRecordset1.EOF Then
								dConceptAntiquityAmount = oRecordset1.Fields("ConceptAmount").Value
							End If
							oRecordset1.Close
						End If
						dAnotherConceptAmount = 0
						Call GetAnotherConceptAmount(iEmployeeID, lStartDate, lEndDate, dAnotherConceptAmount)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesServicesSheetAmounts Set Concept06Amount=" & dConceptAntiquityAmount & ", AnotherConceptAmount =" & dAnotherConceptAmount & " Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
'-->> Ojo
			asEmployeesNumbers = Split(sEmployeesNumbers, NUMERIC_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesNumbers)
				sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & asEmployeesNumbers(iIndex)
				sDocumentName = Server.MapPath(sFileName)
				sHeaderContentsForEmployee = sHeaderContents
				'sQuery = "Select * from Employees, EmployeesExtraInfo Where (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				sQuery = "Select * from Employees Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						'lErrorNumber = AppendTextToFile(sDocumentName, "{\pard{\fs15 Nombre completo:\line}\par}", sErrorDescription)
						sEmployeeName = CStr(oRecordset.Fields("EmployeeName").Value)
						sEmployeeName = sEmployeeName & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sEmployeeName = sEmployeeName & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
						End If
						sHeaderContentsForEmployee = Replace(sHeaderContentsForEmployee, "<DATE />", DisplayNumericDateFromSerialNumber(Left(sDate,8)))
						sHeaderContentsForEmployee = Replace(sHeaderContentsForEmployee, "<EMPLOYEE_NAME />", sEmployeeName)
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", sHeaderContentsForEmployee, sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs18 1. Datos del trabajador:\line}\par}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs15 Nombre completo:}\par}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx2500 \cellx5000 \cellx8000 \cellx11000 \cellx13000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs18" & oRecordset.Fields("EmployeeLastName").Value & " \intbl\cell", sErrorDescription)
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("EmployeeLastName2").Value & " \intbl\cell", sErrorDescription)
							Else
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & " " & " \intbl\cell", sErrorDescription)
							End If
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("EmployeeName").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("RFC").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("CURP").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx14000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc _____________________________________________________________________________________________________________________________________________ \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx2500 \cellx5000 \cellx8000 \cellx11000 \cellx13000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Apellido paterno \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Apellido materno \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Nombre(s) \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc RFC \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc CURP \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard \line \par}", sErrorDescription)
						Next
					End If
					oRecordset.Close
				End If
				sQuery = "Select * From EmployeesExtraInfo Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							'lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs15 Dirección completa:\line}\par}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs15 Dirección completa:}\par}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs18" & oRecordset.Fields("EmployeeAddress").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("EmployeeCity").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("EmployeeZipCode").Value & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx14000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc _____________________________________________________________________________________________________________________________________________ \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Dirección \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Ciudad \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Código Postal \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					Else
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs15 Dirección completa:}\par}", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs18" & "NA" & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "NA" & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "NA" & " \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx14000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc _____________________________________________________________________________________________________________________________________________ \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Dirección \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Ciudad \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Código Postal \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					End If
					oRecordset.Close
				End If
				For iDocIndex = 0 To UBound(asDocumentSuffix)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard \line \par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs18 2. Periódo de aportaciones al fondo del ISSSTE \line}\par}", sErrorDescription)
				Next
				sQuery = "Select Top 1 * from EmployeesServicesSheet Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") Order By StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
						sStartDate = DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
						sReasonName = CStr(oRecordset.Fields("ReasonName").Value)
					End If
					oRecordset.Close
				End If
				sQuery = "Select Top 1 * from EmployeesServicesSheet Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") Order By StartDate Desc"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Select Case CInt(oRecordset.Fields("StatusID").Value)
							Case 26, 30, 34, 38, 42, 46, 50, 122, 129, 144, 148, 155
								lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
								sEndDate = DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
								sEndReasonName = CStr(oRecordset.Fields("ReasonName").Value)
								lLastHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
							Case Else
								sEndDate = "NA"
								sEndReasonName = CStr(oRecordset.Fields("ReasonName").Value)
								lLastHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
						End Select
					End If
					oRecordset.Close
				End If
				For iDocIndex = 0 To UBound(asDocumentSuffix)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000 \cellx12000", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs14" & "Fecha de ingreso" & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "" & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "Fecha de baja" & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "" & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)

					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx3000 \cellx6000 \cellx9000 \cellx12000", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & sStartDate & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "" & sReasonName & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & sEndDate & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & sEndReasonName & " \intbl\cell", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)

					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard__________________________________________________________________________________________________________\par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard \line \par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs18 3. Motivo y periódo en que ocurrio la(s) baja(s), reingreso(s), y/o suspensión(es) \line}\par}", sErrorDescription)
				Next

				sQuery = "Select * From EmployeesServicesSheetAmounts Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") And (SectionID=0) Order By HistoryCounter, PositionID, StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \fs14 \cellx1500 \cellx3000 \cellx7000 \cellx9000 \cellx11000 \cellx12500 \cellx14000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Fecha Inicio \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Fecha Fin \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Puesto \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Pagaduría \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Sueldo cotizable \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Quinquenios \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Otras percepciones \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
						Next
						Do While Not oRecordset.EOF
							' Pinta registros de aportaciones
							For iDocIndex = 0 To UBound(asDocumentSuffix)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \fs14 \cellx1500 \cellx3000 \cellx7000 \cellx9000 \cellx11000 \cellx12500 \cellx14000", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("PositionName").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "50002" & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("Concept01Amount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("Concept06Amount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("AnotherConceptAmount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							Next
							oRecordset.MoveNext
							If Err.number <> 0 Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then Exit Do
						Loop
						oRecordset.Close
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					Else
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx15000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs18 No se encontraron aportaciones \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					End If
				End If
				For iDocIndex = 0 To UBound(asDocumentSuffix)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard__________________________________________________________________________________________________________\par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard \line \par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs18 4. Observaciones \line}\par}", sErrorDescription)
				Next
				If Len(sComments) > 0 Then
					For iDocIndex = 0 To UBound(asDocumentSuffix)
						lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs14 Numero de empleado: " & asEmployeesNumbers(iIndex) & " " & sComments & "}\par}", sErrorDescription)
					Next
				Else
					For iDocIndex = 0 To UBound(asDocumentSuffix)
						lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs14 Numero de empleado: " & asEmployeesNumbers(iIndex) & " NINGUNA OBSERVACIÓN}\par}", sErrorDescription)
					Next
				End If
				For iDocIndex = 0 To UBound(asDocumentSuffix)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard__________________________________________________________________________________________________________\par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard \line \par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard{\fs18 5. Percepciones que aportaron al fondo del ISSSTE \line}\par}", sErrorDescription)
				Next
				sQuery = "Select * From EmployeesServicesSheetAmounts Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") And (SectionID=1) Order By HistoryCounter, PositionID, StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \fs14 \cellx1500 \cellx3000 \cellx7000 \cellx9000 \cellx11000 \cellx12500 \cellx14000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Fecha Inicio \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Fecha Fin \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Puesto \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Pagaduría \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Sueldo cotizable \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Quinquenios \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc Otras percepciones \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
						Next
						Do While Not oRecordset.EOF
							' Pinta registros de aportaciones
							For iDocIndex = 0 To UBound(asDocumentSuffix)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \fs14 \cellx1500 \cellx3000 \cellx7000 \cellx9000 \cellx11000 \cellx12500 \cellx14000", sErrorDescription)
							Next
							If (CLng(oRecordset.Fields("StartDate").Value) < CLng(oRecordset.Fields("HistoryStartDate").Value)) Then
								For iDocIndex = 0 To UBound(asDocumentSuffix)
									lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("HistoryStartDate").Value)) & " \intbl\cell", sErrorDescription)
								Next
							Else
								For iDocIndex = 0 To UBound(asDocumentSuffix)
									lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & " \intbl\cell", sErrorDescription)
								Next
							End If
							If (CLng(oRecordset.Fields("EndDate").Value) > CLng(oRecordset.Fields("HistoryEndDate").Value)) Then
								For iDocIndex = 0 To UBound(asDocumentSuffix)
									lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("HistoryEndDate").Value)) & " \intbl\cell", sErrorDescription)
								Next
							Else
								If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
									For iDocIndex = 0 To UBound(asDocumentSuffix)
										lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "A la fecha" & " \intbl\cell", sErrorDescription)
									Next
								Else
									For iDocIndex = 0 To UBound(asDocumentSuffix)
										lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & " \intbl\cell", sErrorDescription)
									Next
								End If
							End If
							For iDocIndex = 0 To UBound(asDocumentSuffix)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("PositionName").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & "50002" & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("Concept01Amount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("Concept06Amount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc " & oRecordset.Fields("AnotherConceptAmount").Value & " \intbl\cell", sErrorDescription)
								lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							Next
							oRecordset.MoveNext
							If (Err.number <> 0) Or ((CLng(oRecordset.Fields("EndDate").Value) = lLastHistoryStartDate) And (CLng(oRecordset.Fields("EndDate").Value) = 30000000)) Then Exit Do
						Loop
						oRecordset.Close
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					Else
						For iDocIndex = 0 To UBound(asDocumentSuffix)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\trowd \cellx15000", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\qc \fs18 No se encontraron aportaciones \intbl\cell", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "\row", sErrorDescription)
							lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
						Next
					End If
				End If
				For iDocIndex = 0 To UBound(asDocumentSuffix)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "{\pard__________________________________________________________________________________________________________\par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(sDocumentName & asDocumentSuffix(iDocIndex) & ".rtf", "}", sErrorDescription)
				Next
			Next
			oRecordset.Close
			lErrorNumber = ZipFile(Server.MapPath(sFolderPath), Server.MapPath(sZipFile), sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesDocs Set ReportName='" & sDate & "', bPrinted=1 Where (EmployeeID=" & oRequest("EmployeeNumber").Item & ") And (DocumentDate=" & oRequest("DocumentDate").Item & ")", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Path: Update EmployeesDocs Set ReportName='" & sDate & "', bPrinted=1" & sCondition & ">"
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

	Set oRecordset = Nothing
	BuildReport1203 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1203a(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de créditos a los empleados
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1203a"
	Dim sHeaderContents
	Dim sHeaderContentsForEmployee
	Dim oRecordset
	Dim oRecordset1
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim sQuery
	Dim sCondition
	Dim sCondition2
	Dim sEmployeeName
	Dim sEmployeesNumbers
	Dim asEmployeesNumbers
	Dim sActiveEmployeesStatus
	Dim sCancelEmployeesStatus
	Dim lStartDate
	Dim lEndDate
	Dim sStartDate
	Dim sEndDate
	Dim sReasonName
	Dim sEndReasonName
	Dim iHistoryCount
	Dim lHistoryStartDate
	Dim lHistoryEndDate
	Dim lLastHistoryStartDate
	Dim dConceptAntiquityAmount
	Dim dAnotherConceptAmount
	Dim sComments
	Dim sDocumentTypeCondition
	Dim sDocumentTypeConditionForCV

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asCreditsNames
	Dim asPath
	Dim iCount
	Dim aiCreditsTotals
	Dim aiCreditsGrandTotals
	Dim iIndex
	Dim sCreditShortName
	Dim bFirst
	Dim lTotal
	Dim iMin
	Dim iMax

	sDocumentSuffix = "_01,_02,_03,_04,_05,_06"
	asDocumentSuffix = Split(sDocumentSuffix, ",", -1, vbBinaryCompare)

	sActiveEmployeesStatus = "-1,0,1,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,27,28,29,31,32,33,35,36,37,39,40,41,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,119,120,121,123,124,125,126,127,128,130,131,132,133,134,135,136,137,138,139,140,141,142,143,145,146,147,149,150,151,152,153,154,155,156,157,158"
	sCancelEmployeesStatus = "2,3,4,26,30,34,38,42,118,122,129,144,148,159"

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If Len(oRequest("EmployeeNumber").Item) > 0 Then 
		sCondition = sCondition & " And (EmployeesDocs.EmployeeID=" & oRequest("EmployeeNumber").Item & ")"
	Else
		sCondition = sCondition & " And (EmployeesDocs.EmployeeID=-1)"
	End If
	If (InStr(1, oRequest, "DocumentStart", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "DocumentEnd", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("DocumentStart", "DocumentEnd", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	'If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesDocs.DocumentStart") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesDocs.DocumentEnd") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesDocs.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesDocs.DocumentStart", 1, 1, vbBinaryCompare) & "))"
	If Len(sCondition2) > 0 Then sCondition2 = " And " & Replace(sCondition2, "XXX", "EmployeesDocs.Document")
	If Len(oRequest("DocumentTypeID").Item) > 0 Then
		If CInt(oRequest("DocumentTypeID").Item) = 1 Then
			Dim lDocumentTypeStartDate, lDocumentTypeEndDate
			lDocumentTypeEndDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
			lDocumentTypeStartDate = AddYearsToSerialDate(Left(GetSerialNumberForDate(""), Len("00000000")), -1)
			sDocumentTypeCondition = " And (((EHL.EmployeeDate>=" & lDocumentTypeStartDate & ") And (EHL.EmployeeDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((EHL.EndDate>=" & lDocumentTypeStartDate & ") And (EHL.EndDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((EHL.EndDate>=" & lDocumentTypeStartDate & ") And (EHL.EmployeeDate<=" & lDocumentTypeEndDate & ")))"
			sDocumentTypeConditionForCV = " And (((ConceptsValues.StartDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.StartDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((ConceptsValues.EndDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.EndDate<=" & lDocumentTypeEndDate & "))" & _
							" Or ((ConceptsValues.EndDate>=" & lDocumentTypeStartDate & ") And (ConceptsValues.StartDate<=" & lDocumentTypeEndDate & ")))"
		End If
	End If
	If Len(oRequest("ShowServiceSheet").Item) > 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesDocs.EmployeeID, EmployeesDocs.Comments From EmployeesDocs, Employees Where (EmployeesDocs.EmployeeID=Employees.EmployeeID)" & sCondition & sCondition2 & " Order By EmployeesDocs.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "Seleccione la acción 'Mostrar documento previo' del listado de Hojas únicas de servicio solicitadas, que se muestran en la sección 2."
	End If
	If lErrorNumber = 0 Then
		sEmployeesNumbers = ""
		If Not oRecordset.EOF Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesServicesSheet", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesServicesSheetAmounts", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1203.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_YEAR />", Year(Date()))
			Do While Not oRecordset.EOF
				sEmployeesNumbers = sEmployeesNumbers & CStr(oRecordset.Fields("EmployeeID").Value) & NUMERIC_SEPARATOR
				sComments = CStr(oRecordset.Fields("Comments").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			sEmployeesNumbers = Left(sEmployeesNumbers, (Len(sEmployeesNumbers) - Len(",")))
			oRecordset.Close
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen empleados registrados para la entrega de hoja única de servicios para los criterios indicados."
		End If
	End If
	If lErrorNumber = 0 Then
		oStartDate = Now()
		sErrorDescription = "No se pudieron obtener los registros del empleado."
		sCondition2 = " And (Employees.EmployeeID IN (" & sEmployeesNumbers & "))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheet Select Distinct EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID, 0 As CashOfficer, 0 As Concept01Amount, 0 As Concept06Amount, 0 As AnotherConceptAmount From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, StatusEmployees, Reasons, Companies, EmployeeTypes, Employees Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.PositionID=Positions.PositionID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.EmployeeID=Employees.EmployeeID) And (EHL.EmployeeDate<=EHL.EndDate)" & sCondition2 & sDocumentTypeCondition & " Group by EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: " & "Insert Into EmployeesServicesSheet Select Distinct EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID, 0 As CashOfficer, 0 As Concept01Amount, 0 As Concept06Amount, 0 As AnotherConceptAmount From EmployeesHistoryList EHL, Zones, Areas, Areas As PaymentCenters, Positions, StatusEmployees, Reasons, Companies, EmployeeTypes, Employees Where (EHL.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EHL.CompanyID=Companies.CompanyID) And (EHL.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EHL.PaymentCenterID=PaymentCenters.AreaID) And (EHL.PositionID=Positions.PositionID) And (EHL.StatusID=StatusEmployees.StatusID) And (EHL.ReasonID=Reasons.ReasonID) And (EHL.EmployeeID=Employees.EmployeeID) And (EHL.EmployeeDate<=EHL.EndDate)" & sCondition2 & sDocumentTypeCondition & " Group by EHL.EmployeeID, EHL.EmployeeDate, EHL.EndDate, EHL.PositionID, PositionName, EHL.StatusID, StatusName, EHL.ReasonID, ReasonName, EHL.CompanyID, PaymentCenters.EconomicZoneID, EHL.LevelID, EHL.GroupGradeLevelID, EHL.IntegrationID, EHL.ClassificationID, EHL.WorkingHours, EHL.EmployeeTypeID Order By EHL.EmployeeID, EmployeeDate Desc, EHL.EndDate Desc" & " -->" & vbNewLine
		If lErrorNumber = 0 Then
'-->> Ojo
			sDate = GetSerialNumberForDate("")
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheet Where (StatusID IN(" & sActiveEmployeesStatus & ")) Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			iHistoryCount = 0
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						iHistoryCount = iHistoryCount + 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 1 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 1 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV & " -->" & vbNewLine
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheet Where (StatusID IN(" & sCancelEmployeesStatus & ")) Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			'iHistoryCount = 0
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						iHistoryCount = iHistoryCount + 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 0 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesServicesSheetAmounts Select Distinct '" & oRecordset.Fields("EmployeeID").Value & "' As EmployeeID, StartDate, EndDate, '" & oRecordset.Fields("PositionID").Value & "' As PositionID, " & iHistoryCount & " As HistoryCount, " & oRecordset.Fields("EmployeeTypeID").Value & " As EmployeeTypeID, " & oRecordset.Fields("StartDate").Value & " As HistoryStartDate, " & oRecordset.Fields("EndDate").Value & " As HistoryEndDate, '" & oRecordset.Fields("PositionName").Value & "' As PositionName, 0 As SectionID, '50002' As CashOfficer, ConceptAmount, 0 As Concept06Amount, 0 As AnotherConceptAmount From ConceptsValues Where (PositionID=" & oRecordset.Fields("PositionID").Value & ") And ((CompanyID=" & oRecordset.Fields("CompanyID").Value & ") Or (CompanyID=-1)) And ((EconomicZoneID=" & oRecordset.Fields("EconomicZoneID").Value & ") Or (EconomicZoneID=0)) And ((LevelID=" & oRecordset.Fields("LevelID").Value & ") Or (LevelID=0)) And (IntegrationID=" & oRecordset.Fields("IntegrationID").Value & ") And (ClassificationID=" & oRecordset.Fields("ClassificationID").Value & ") And ((GroupGradeLevelID=" & oRecordset.Fields("GroupGradeLevelID").Value & ") Or (GroupGradeLevelID=-1)) And ((WorkingHours=" & oRecordset.Fields("WorkingHours").Value & ") Or (WorkingHours=-1)) And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) And (ConceptID=1)" & sDocumentTypeConditionForCV & " -->" & vbNewLine
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesServicesSheetAmounts Order By EmployeeID, StartDate", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						lStartDate = oRecordset.Fields("StartDate").Value
						lEndDate = oRecordset.Fields("EndDate").Value
						aEmployeeComponent(N_ID_EMPLOYEE) = oRecordset.Fields("EmployeeID").Value
						lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, lStartDate, sEmployeeAntiquity, lAntiquityYears, lAntiquityMonths, lAntiquityDays, sErrorDescription)
						lAntiquityYears = lAntiquityYears + lAntiquityMonths/12
						'Call GetConceptAntiquityAmount(lAntiquity, dConceptAntiquityAmount)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from ConceptsValues Where (ConceptID=6) And (Antiquity2ID<=" & lAntiquityYears & ") And (EmployeeTypeID=" & oRecordset.Fields("EmployeeTypeID").Value & ") And (((StartDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ")) Or ((EndDate>=" & oRecordset.Fields("StartDate").Value & ") And (StartDate<=" & oRecordset.Fields("EndDate").Value & "))) Order By Antiquity2ID Desc", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						If lErrorNumber = 0 Then
							If Not oRecordset1.EOF Then
								dConceptAntiquityAmount = oRecordset1.Fields("ConceptAmount").Value
							End If
							oRecordset1.Close
						End If
						dAnotherConceptAmount = 0
						Call GetAnotherConceptAmount(iEmployeeID, lStartDate, lEndDate, dAnotherConceptAmount)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesServicesSheetAmounts Set Concept06Amount=" & dConceptAntiquityAmount & ", AnotherConceptAmount =" & dAnotherConceptAmount & " Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
'-->> Ojo
			asEmployeesNumbers = Split(sEmployeesNumbers, NUMERIC_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesNumbers)
				'sFileName = sFolderPath & "\User_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aReportsComponent(N_ID_REPORTS) & "_" & asEmployeesNumbers(iIndex)
				'sDocumentName = Server.MapPath(sFileName & ".rtf")
				'sHeaderContentsForEmployee = sHeaderContents
				'sQuery = "Select * from Employees, EmployeesExtraInfo Where (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				sQuery = "Select * from Employees Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sContents = Replace(sContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)))
						sContents = Replace(sContents, "<EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sContents = Replace(sContents, "<EMPLOYEE_LAST_NAME_2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)))
						Else
							sContents = Replace(sContents, "<EMPLOYEE_LAST_NAME_2 />", " ")
						End If
						sContents = Replace(sContents, "<EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
						sContents = Replace(sContents, "<EMPLOYEE_CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
					End If
					oRecordset.Close
				End If
				sQuery = "Select * From EmployeesExtraInfo Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sContents = Replace(sContents, "<EMPLOYEE_ADDRES />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeAddress").Value)))
						sContents = Replace(sContents, "<EMPLOYEE_CITY />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeCity").Value)))
						sContents = Replace(sContents, "<EMPLOYEE_ZIP_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeZipCode").Value)))
					Else
						sContents = Replace(sContents, "<EMPLOYEE_ADDRES />", "NA")
						sContents = Replace(sContents, "<EMPLOYEE_CITY />", "NA")
						sContents = Replace(sContents, "<EMPLOYEE_ZIP_CODE />", "NA")
					End If
					oRecordset.Close
				End If
				'lErrorNumber = AppendTextToFile(sDocumentName, "{\pard \line \par}", sErrorDescription)
				'lErrorNumber = AppendTextToFile(sDocumentName, "{\pard{\fs18 2. Periódo de aportaciones al fondo del ISSSTE \line}\par}", sErrorDescription)

				sQuery = "Select Top 1 * from EmployeesServicesSheet Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") Order By StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
						sStartDate = DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
						sReasonName = CStr(oRecordset.Fields("ReasonName").Value)
					End If
					oRecordset.Close
				End If
				sQuery = "Select Top 1 * from EmployeesServicesSheet Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") Order By StartDate Desc"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Select Case CInt(oRecordset.Fields("StatusID").Value)
							Case 26, 30, 34, 38, 42, 46, 50, 122, 129, 144, 148, 155
								lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
								sEndDate = DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
								sEndReasonName = CStr(oRecordset.Fields("ReasonName").Value)
								lLastHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
							Case Else
								sEndDate = "NA"
								sEndReasonName = CStr(oRecordset.Fields("ReasonName").Value)
								lLastHistoryStartDate = CLng(oRecordset.Fields("StartDate").Value)
						End Select
					End If
					oRecordset.Close
				End If
				sContents = Replace(sContents, "<START_DATE />", CleanStringForHTML(sStartDate))
				sContents = Replace(sContents, "<REASON_START />", CleanStringForHTML(sReasonName))
				sContents = Replace(sContents, "<END_DATE />", CleanStringForHTML(sEndDate))
				sContents = Replace(sContents, "<REASON_END />", CleanStringForHTML(sEndReasonName))
				Response.Write sContents
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"				
				sQuery = "Select * From EmployeesServicesSheetAmounts Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") And (SectionID=0) Order By HistoryCounter, PositionID, StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD>Fecha Inicio</TD>"
							Response.Write "<TD>Fecha Fin</TD>"
							Response.Write "<TD>Puesto</TD>"
							Response.Write "<TD>Pagaduría</TD>"
							Response.Write "<TD>Sueldo cotizable</TD>"
							Response.Write "<TD>Quinquenios</TD>"
							Response.Write "<TD>Otras percepciones</TD>"
						Response.Write "</TR>"
						Do While Not oRecordset.EOF

							Response.Write "<TR>"
								Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</TD>"
								Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("PositionName").Value & "</TD>"
								Response.Write "<TD>" & "50002" & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("Concept01Amount").Value & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("Concept06Amount").Value & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("AnotherConceptAmount").Value & "</TD>"
							Response.Write "</TR>"
						
							oRecordset.MoveNext
							If Err.number <> 0 Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then Exit Do
						Loop
						oRecordset.Close
					Else
						Response.Write "<TR>"
							Response.Write "<TD>" & "NO SE ENCONTRARON REGISTROS" & "</TD>"
						Response.Write "</TR>"
					End If
				End If
				Response.Write "</TABLE><BR />"
				Response.Write "<LINE />"
				Response.Write "<B>4. Observaciones</B><BR /><BR />"
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD>"
					If Len(sComments) > 0 Then
						Response.Write sComments
					Else
						Response.Write "NA"
					End If
				Response.Write "</TD></TR></TABLE><BR />"
				Response.Write "<B>5. Percepciones que aportaron al fondo del ISSSTE</B><BR /><BR />"				
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				sQuery = "Select * From EmployeesServicesSheetAmounts Where (EmployeeID=" & asEmployeesNumbers(iIndex) & ") And (SectionID=1) Order By HistoryCounter, PositionID, StartDate"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD>Fecha Inicio</TD>"
							Response.Write "<TD>Fecha Fin</TD>"
							Response.Write "<TD>Puesto</TD>"
							Response.Write "<TD>Pagaduría</TD>"
							Response.Write "<TD>Sueldo cotizable</TD>"
							Response.Write "<TD>Quinquenios</TD>"
							Response.Write "<TD>Otras percepciones</TD>"
						Response.Write "</TR>"
						Do While Not oRecordset.EOF
							' Pinta registros de aportaciones
							Response.Write "<TR>"
								If (CLng(oRecordset.Fields("StartDate").Value) < CLng(oRecordset.Fields("HistoryStartDate").Value)) Then
									Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("HistoryStartDate").Value)) & "</TD>"								
								Else
									Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</TD>"
								End If
								If (CLng(oRecordset.Fields("EndDate").Value) > CLng(oRecordset.Fields("HistoryEndDate").Value)) Then
									Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("HistoryEndDate").Value)) & "</TD>"
								Else
									If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
										Response.Write "<TD>" & "A la fecha" & "</TD>"
									Else
										Response.Write "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</TD>"
									End If
								End If
								Response.Write "<TD>" & oRecordset.Fields("PositionName").Value & "</TD>"
								Response.Write "<TD>" & "50002" & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("Concept01Amount").Value & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("Concept06Amount").Value & "</TD>"
								Response.Write "<TD>" & oRecordset.Fields("AnotherConceptAmount").Value & "</TD>"
							Response.Write "</TR>"
							oRecordset.MoveNext
							If (Err.number <> 0) Or ((CLng(oRecordset.Fields("EndDate").Value) = lLastHistoryStartDate) And (CLng(oRecordset.Fields("EndDate").Value) = 30000000)) Then Exit Do
						Loop
						oRecordset.Close
					Else
						Response.Write "<TR>"
							Response.Write "<TD>" & "NO SE ENCONTRARON REGISTROS" & "</TD>"
						Response.Write "</TR>"
					End If
				End If
				Response.Write "</TABLE><BR />"
				Response.Write "<LINE />"
			Next
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1203a = lErrorNumber
	Err.Clear
End Function

Function BuildReport1202Sp(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal con conceptos. Reporte basado en la hoja 001221 
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1202Sp"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim oCompaniesRecordset
	Dim sCurrentRecords
	Dim sTempRecords
	Dim asColumnsTitles
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sHeaderContents
	Dim asConceptTitle
	Dim bEmpty
	Dim lTotalEmployees
	Dim dTotalAmount
	Dim iCompany
	Dim lTotalPayolls
	Dim lStartDate
	Dim lRecordDate

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If (InStr(1, sCondition, "Concepts.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "Concepts.", "Percepciones.")
	End If

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	lTotalEmployees = 0
	dTotalAmount = 0
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		If lErrorNumber = 0 Then
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()		
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1024.htm"), sErrorDescription)
			If Len(sHeaderContents) > 0 Then
				asConceptTitle = Split(aReportTitle(L_CONCEPT_ID_FLAGS), ";")
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NAME />", asConceptTitle(1))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NUMBER />", asConceptTitle(0))
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
			End If
			lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
			If Not (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyName From Companies Where (ParentID>=0) And (EndDate=30000000) Order By CompanyShortName", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCompaniesRecordset)
				If lErrorNumber = 0 Then
					If Not oCompaniesRecordset.EOF Then
						Do While Not oCompaniesRecordset.EOF
							iCompany = CInt(oCompaniesRecordset.Fields("CompanyID").Value)
							sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Credits.StartDate, Percepciones.RecordDate, Percepciones.ConceptAmount, EmployeeTypes.EmployeeTypeShortName, Credits.StartDate From Employees, EmployeeTypes, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas, Credits Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.EmployeeTypeID = EmployeeTypes.EmployeeTypeID) And (Employees.LevelID = Levels.LevelID) And (Employees.EmployeeID = Credits.EmployeeID) And (Credits.CreditTypeID = 60) And(Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								lTotalEmployees = 0
								dTotalAmount = 0
								If Not oRecordset.EOF Then
									bEmpty = False
									sRowContents = "<BR /><B>EMPRESA:" & Cstr(oRecordset.Fields("CompanyName").Value) &  "</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "<TD>TIPO</TD>"
									sRowContents = sRowContents & "<TD>No.CUOTAS</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
									    lStartDate = CLng(oRecordset.Fields("StartDate").Value)
										lRecordDate = CLng(oRecordset.Fields("RecordDate").Value)
										lTotalPayrolls = GetTotalPayrolls(lStartDate, lRecordDate )
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										Else
											sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
										End If
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value)) & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">" & lTotalPayrolls & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " FUNCIONARIOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR /><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " And Employees.CompanyID=" & iCompany & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lTotalEmployees = 0
									dTotalAmount = 0
									bEmpty = False
									sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>No.EMP.</TD>"
									sRowContents = sRowContents & "<TD>NOMBRE</TD>"
									sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
									sRowContents = sRowContents & "<TD>RFC</TD>"
									sRowContents = sRowContents & "<TD>PUESTO</TD>"
									sRowContents = sRowContents & "<TD>N/SN</TD>"
									sRowContents = sRowContents & "<TD>MONTO</TD>"
									sRowContents = sRowContents & "</TR></FONT>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									Do While Not oRecordset.EOF
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										Else
											sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
										End If
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
										sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
										sRowContents = sRowContents & "</TD>"
										sRowContents = sRowContents & "</TR></FONT>"
										lTotalEmployees = lTotalEmployees + 1
										dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
										oRecordset.MoveNext
										If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									sRowContents = "</TABLE>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />TOTAL " & CStr(oCompaniesRecordset.Fields("CompanyName").Value) &  " OPERATIVOS: "  & lTotalEmployees
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True) & "<BR />"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							End If
							oCompaniesRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oCompaniesRecordset.Close
					End If
				End If
			Else
				sCondition = Replace(sCondition, "Companies.", "Employees.")
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (1)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						bEmpty = False
						lTotalEmployees = 0
						dTotalAmount = 0
						sRowContents = "<BR /><B>Empresa: " & aReportTitle(L_COMPANY_FLAGS) & "</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>FUNCIONARIOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO FUNCIONARIOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Areas.AreaCode, Employees.RFC, Positions.PositionShortName, Levels.LevelShortName, Companies.CompanyName, Percepciones.ConceptAmount From Employees, Companies, Payroll_" & lPayrollID & " As Percepciones, Jobs, Positions, Levels, Areas Where (Percepciones.EmployeeID = Employees.EmployeeID) And (Employees.EmployeeTypeID In (0,2,3,4,5,6)) And (Employees.CompanyID = Companies.CompanyID) And (Employees.JobID = Jobs.JobID) And (Areas.AreaID = Jobs.AreaID) And (Employees.LevelID = Levels.LevelID) And (Positions.PositionID = Jobs.PositionID) " & sCondition & " Order By Employees.EmployeeTypeID, Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						bEmpty = False
						sRowContents = "<BR /><B>OPERATIVOS</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>No.EMP.</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE</TD>"
						sRowContents = sRowContents & "<TD>ADSCRIP.</TD>"
						sRowContents = sRowContents & "<TD>RFC</TD>"
						sRowContents = sRowContents & "<TD>PUESTO</TD>"
						sRowContents = sRowContents & "<TD>N/SN</TD>"
						sRowContents = sRowContents & "<TD>MONTO</TD>"
						sRowContents = sRowContents & "</TR></FONT>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lTotalEmployees = 0
						dTotalAmount = 0
						Do While Not oRecordset.EOF
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							Else
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							End If
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">"
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT"">"
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
							sRowContents = sRowContents & "</TD>"
							sRowContents = sRowContents & "</TR></FONT>"
							lTotalEmployees = lTotalEmployees + 1
							dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR />TOTAL " & aReportTitle(L_COMPANY_FLAGS) &  " FUNCIONARIOS: "  & lTotalEmployees
						sRowContents = "<BR />MONTO OPERATIVOS $ " & FormatNumber(dTotalAmount, 2, True, False, True)
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				End If
			End If
			If Not bEmpty Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1202Sp = lErrorNumber
	Err.Clear
End Function

Function BuildReport1207(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Constancia de servicio activo
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1207"
	Dim sContents
	Dim sCondition
	Dim lPayrollID
	Dim sTemp
	Dim dTotal
	Dim oRecordset
	Dim lErrorNumber

	Call GetNameFromTable(oADODBConnection, "LastClosedPayrollID", "-1", "", "", lPayrollID, sErrorDescription)
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = sCondition & " And (ConceptID In (1,6,7,8,4,5))"
	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1207.htm"), sErrorDescription)
	If Len(sContents) > 0 Then
		sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		sContents = Replace(sContents, "<CURRENT_YEAR />", Year(Date()))
		sContents = Replace(sContents, "<CURRENT_DAY_AS_TEXT />", FormatNumberAsText(Day(Date()), False))
		sContents = Replace(sContents, "<CURRENT_MONTH />", asMonthNames_es(Month(Date())))
		sContents = Replace(sContents, "<CURRENT_YEAR_AS_TEXT />", FormatNumberAsText(Year(Date()), False))
		sContents = Replace(sContents, "<DOCUMENT_NUMBER />", CleanStringForHTML(oRequest("DocumentNumber").Item))
		sContents = Replace(sContents, "<COMMENTS />", CleanStringForHTML(oRequest("Comments").Item))
		sContents = Replace(sContents, "<PURPOSE />", CleanStringForHTML(oRequest("Purpose").Item))
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, Employees.StartDate, Areas.AreaCode, Areas.AreaName, PositionName, PositionTypeName, ConceptID, ConceptAmount From Employees, Payroll_" & lPayrollID & ", Jobs, Positions, PositionTypes, Areas Where (Employees.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) " & sCondition, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sTemp = CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
				sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(sTemp))
				sContents = Replace(sContents, "<EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
				sContents = Replace(sContents, "<EMPLOYEE_CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
				sContents = Replace(sContents, "<EMPLOYEE_START_DATE />", DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1))
				sContents = Replace(sContents, "<POSITION_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)))
				sContents = Replace(sContents, "<POSITION_TYPE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeName").Value)))
				sContents = Replace(sContents, "<AREA_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)))
				sContents = Replace(sContents, "<AREA_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)))
				sContents = Replace(sContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
				If bForExport Then
					sContents = Replace(sContents, "<LINE />", "")
					sContents = Replace(sContents, "<LINE2 />", "<TR><TD COLSPAN=""5"">&nbsp;<BR /><BR /><BR /><BR /><BR /></TD></TR>")
				Else
					sContents = Replace(sContents, "<LINE />", "<TR><TD BGCOLOR=""#000000"" COLSPAN=""5""><IMG SRC=""<SYSTEM_URL />Templates/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR><TR><TD COLSPAN=""5""><IMG SRC=""<SYSTEM_URL />Templates/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"" /></TD></TR>")
					sContents = Replace(sContents, "<LINE2 />", "<TR><TD BGCOLOR=""#000000"" COLSPAN=""5""><IMG SRC=""<SYSTEM_URL />Templates/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR><TR><TD COLSPAN=""5""><IMG SRC=""<SYSTEM_URL />Templates/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"" /></TD></TR>")
				End If
				sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				dTotal = 0
				Do While Not oRecordset.EOF
					sContents = Replace(sContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value) * 2, 2, True, False, True))
					dTotal = dTotal + (CDbl(oRecordset.Fields("ConceptAmount").Value) * 2)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sContents = Replace(sContents, "<CONCEPT_1 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_6 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_7 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_8 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_4 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_5 />", "-------")
				sContents = Replace(sContents, "<CONCEPT_99 />", "-------")
				sContents = Replace(sContents, "<TOTAL />", FormatNumber(dTotal, 2, True, False, True))
			End If
			oRecordset.Close
		End If
		Response.Write sContents
	End If

	Set oRecordset = Nothing
	BuildReport1207 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1208(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Constancia de descuento
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1208"
	Dim sTemplate
	Dim sContents
	Dim sCondition
	Dim lStartDate
	Dim lEndDate
	Dim asPayrolls
	Dim sTemp
	Dim iIndex
	Dim jIndex
	Dim iCounter
	Dim oRecordset
	Dim lErrorNumber

	lStartDate = 20080101
	If (CInt(oRequest("StartYear").Item) > 0) And (CInt(oRequest("StartMonth").Item) > 0) And (CInt(oRequest("StartDay").Item) > 0) Then lStartDate = CLng(oRequest("StartYear").Item & oRequest("StartMonth").Item & oRequest("StartDay").Item)
	lEndDate = 20101231
	If (CInt(oRequest("EndYear").Item) > 0) And (CInt(oRequest("EndMonth").Item) > 0) And (CInt(oRequest("EndDay").Item) > 0) Then lEndDate = CLng(oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item)
	If lStartDate > lEndDate Then
		sTemp = lStartDate
		lStartDate = lEndDate
		lEndDate = sTemp
	End If
	If Len(oRequest("EmployeeNumber").Item) > 0 Then
		sCondition = " And (Employees.EmployeeID=" & oRequest("EmployeeNumber").Item & ")"
	Else
		sCondition = " And (Employees.EmployeeID=" & oRequest("EmployeeID").Item & ")"
	End If
	sTemplate = GetFileContents(Server.MapPath("Templates\HeaderForReport_1208.htm"), sErrorDescription)
	If Len(sTemplate) > 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Areas.AreaCode, Areas.AreaName From Employees, Jobs, Areas Where (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) " & sCondition, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sTemplate = Replace(sTemplate, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sTemplate = Replace(sTemplate, "<CURRENT_YEAR />", Year(Date()))
				sTemplate = Replace(sTemplate, "<CURRENT_DAY_AS_TEXT />", UCase(FormatNumberAsText(Day(Date()), False)))
				sTemplate = Replace(sTemplate, "<CURRENT_MONTH />", UCase(asMonthNames_es(Month(Date()))))
				sTemplate = Replace(sTemplate, "<CURRENT_YEAR_AS_TEXT />", UCase(FormatNumberAsText(Year(Date()), False)))
				sTemp = CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
				sTemplate = Replace(sTemplate, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(sTemp))
				sTemplate = Replace(sTemplate, "<EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
				sTemplate = Replace(sTemplate, "<AREA_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)))
				sTemplate = Replace(sTemplate, "<AREA_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)))
				sTemplate = Replace(sTemplate, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
			End If
			oRecordset.Close

			sContents = sTemplate
			asPayrolls = ""
			sTemp = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
				iCounter = 0
				For iIndex = CInt(Left(lStartDate, Len("YYYY"))) To CInt(Left(lEndDate, Len("YYYY")))
					asPayrolls = asPayrolls & iIndex & ","
					sTemp = sTemp & "<TD VALIGN=""TOP""><TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
						sTemp = sTemp & "<TR>"
							sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2"">MES-AÑO/ " & iIndex & "</FONT></TD>"
							sTemp = sTemp & "<TD>&nbsp;&nbsp;&nbsp;</TD>"
							sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2"">1a. QNA.</FONT></TD>"
							sTemp = sTemp & "<TD>&nbsp;&nbsp;&nbsp;</TD>"
							sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2"">2a. QNA.</FONT></TD>"
						sTemp = sTemp & "</TR>"
						For jIndex = 1 To 12
							sTemp = sTemp & "<TR>"
								sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2"">" & asMonthNames_es(jIndex) & "</FONT></TD>"
								sTemp = sTemp & "<TD>&nbsp;&nbsp;&nbsp;</TD>"
								sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2""><" & iIndex & Right(("0" & jIndex), Len("00")) & "01 /></FONT></TD>"
								sTemp = sTemp & "<TD>&nbsp;&nbsp;&nbsp;</TD>"
								sTemp = sTemp & "<TD><FONT FACE=""Arial"" SIZE=""2""><" & iIndex & Right(("0" & jIndex), Len("00")) & "02 /></FONT></TD>"
							sTemp = sTemp & "</TR>"
						Next
					sTemp = sTemp & "</TABLE></TD>"
					sTemp = sTemp & "<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
					iCounter = iCounter + 1
					If (iCounter Mod 2) = 0 Then
						sTemp = sTemp & "</TR></TABLE>"
						sContents = Replace(sContents, "<PAYROLLS />", sTemp)
						If iIndex < CInt(Left(lEndDate, Len("YYYY"))) Then sContents = sContents & sTemplate
						sTemp = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
					End If
				Next
			sTemp = sTemp & "</TR></TABLE>"
			sContents = Replace(sContents, "<PAYROLLS />", sTemp)

			sCondition = "(ConceptID=" & oRequest("ConceptID").Item & ") And (RecordDate>=" & lStartDate & ") And (RecordDate<=" & lEndDate & ") " & Replace(sCondition, "Employees.", "Payroll_YYYY.")
			asPayrolls = Split(asPayrolls, ",")
			For iIndex = 0 To UBound(asPayrolls) - 1
				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordDate, ConceptAmount From Payroll_" & asPayrolls(iIndex) & " Where " & Replace(sCondition, "YYYY", asPayrolls(iIndex)), "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Int(Right(CStr(oRecordset.Fields("RecordDate").Value), Len("DD"))) > 15 Then
							sContents = Replace(sContents, "<" & Left(CStr(oRecordset.Fields("RecordDate").Value), Len("YYYYMM")) & "02 />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
						Else
							sContents = Replace(sContents, "<" & Left(CStr(oRecordset.Fields("RecordDate").Value), Len("YYYYMM")) & "01 />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			Next
			lErrorNumber = 0
			For iIndex = 0 To UBound(asPayrolls) - 1
				For jIndex = 1 To 12
					sContents = Replace(sContents, "<" & asPayrolls(iIndex) & Right(("0" & jIndex), Len("00")) & "01 />", "--------")
					sContents = Replace(sContents, "<" & asPayrolls(iIndex) & Right(("0" & jIndex), Len("00")) & "02 />", "--------")
				Next
			Next
		End If
		Response.Write sContents
	End If

	Set oRecordset = Nothing
	BuildReport1208 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1209(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Listado de registros de empleados para hacer 
'         revisión de pagos sobre nóminas
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1209"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim sCondition
	Dim sCondition2

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesRevisions.AddDate") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesRevisions.AddDate") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesRevisions.AddDate", 1, 1, vbBinaryCompare), "XXX", "EmployeesRevisions.AddDate", 1, 1, vbBinaryCompare) & "))"

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros cargados del archivo de terceros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesRevisions.*, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As FullName, UserName, UserLastName From EmployeesRevisions, Employees, Users Where (EmployeesRevisions.EmployeeID=Employees.EmployeeID) And (EmployeesRevisions.UserID=Users.UserID)" & sCondition & sCondition2 & " Order By PayrollID Desc", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReports.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<TITLE />", CleanStringForHTML("Listado de empleados para revisión de nóminas"))
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Q.Aplicación</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Q.Inicio de Revisión</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Usuario que capturo</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">F.Registro</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Observaciones</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("FullName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(oRecordset.Fields("PayrollID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(oRecordset.Fields("StartPayrollID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UserName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(oRecordset.Fields("AddDate").Value) & "</FONT></TD>"
						If (Not IsNull(oRecordset.Fields("Comments").Value)) And (Len(oRecordset.Fields("Comments").Value) > 0) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("Comments").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1209 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1210(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de horas extras y primas dominicales para el personal
'         filtrado por número de empleado, áreas y período específico
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1210"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sCondition
	Dim sCondition2
	Dim sQuery

	sQuery = "Select EmployeesAbsencesLKP.EmployeeID, Employees.EmployeeNumber, EmployeesAbsencesLKP.AbsenceID, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesAbsencesLKP.AppliedDate," & _
			 " EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.RegistrationDate, EmployeesAbsencesLKP.DocumentNumber, EmployeesAbsencesLKP.AbsenceHours, A.AbsenceShortName, A.AbsenceName," & _
			 " J.JustificationShortName, EmployeesAbsencesLKP.Reasons, EmployeesAbsencesLKP.Removed, EmployeesAbsencesLKP.JustificationID As AbsenceJustified, EmployeesAbsencesLKP.Active," & _
			 " A.IsJustified, A.JustificationID As WithJustification, Users.UserLastName + ' ' + Users.UserName As UserFullName" & _
			 " From Employees, EmployeesAbsencesLKP, Absences As A, Justifications As J, Users, Areas, Areas As PaymentCenters, Jobs" & _
			 " Where (Employees.EmployeeID = EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.JustificationID = J.JustificationID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID = A.AbsenceID) And (EmployeesAbsencesLKP.AddUserID=Users.UserID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID)"

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "OcurredDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("OcurredDate", "EndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.Ocurred") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesAbsencesLKP.Ocurred", 1, 1, vbBinaryCompare) & "))"
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	Call DisplayTimeStamp("START: CONSULTA. " & sQuery & sCondition & sCondition2)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition & sCondition2, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1210.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido materno</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de ocurrencia</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de término</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Aplicada</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Q. de aplicación</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. de documento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Justificación</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del usuario</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
						If CLng(oRecordset.Fields("Active").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NO</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">SI</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("JustificationShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value)) & "</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1210 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1211(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the employee's grades
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1211"
	Dim sCondition
	Dim lYearID
	Dim oRecordset
	Dim sTemp
	Dim dAmount
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "EmployeesHistoryList."), "PaymentCenters.", "EmployeesHistoryList.")
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If

	lYearID = oRequest("YearID").Item

	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryList.JobID, LevelShortName, EmployeesHistoryList.WorkingHours, Areas2.EconomicZoneID, EmployeeGrade, EmployeesHistoryList.CompanyID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.ClassificationID, EmployeesHistoryList.GroupGradeLevelID, EmployeesHistoryList.IntegrationID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.LevelID, EmployeesHistoryList.PositionID, Payroll_" & lYearID & ".ConceptAmount, EmployeesGrades.GradePercentage From Payroll_" & lYearID & ", Employees, EmployeesChangesLKP, EmployeesHistoryList, EmployeesGrades, Levels, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lYearID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=EmployeesGrades.PayrollID) And (EmployeesHistoryList.EmployeeID=EmployeesGrades.EmployeeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Levels.StartDate<=EmployeesGrades.PayrollID) And (Levels.EndDate>=EmployeesGrades.PayrollID) And (Areas1.StartDate<=EmployeesGrades.PayrollID) And (Areas1.EndDate>=EmployeesGrades.PayrollID) And (Areas2.StartDate<=EmployeesGrades.PayrollID) And (Areas2.EndDate>=EmployeesGrades.PayrollID) And (EmployeesGrades.StartDate<=" & lYearID & "0101) And (EmployeesGrades.EndDate>=" & lYearID & "1231) And (Payroll_2011.RecordDate=EmployeesGrades.PayrollID) And (Payroll_2011.RecordID=EmployeesGrades.PayrollID) And (Payroll_" & lYearID & ".ConceptID=1) And (EmployeesGrades.Active=1) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryList.JobID, LevelShortName, EmployeesHistoryList.WorkingHours, Areas2.EconomicZoneID, EmployeeGrade, EmployeesHistoryList.CompanyID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.ClassificationID, EmployeesHistoryList.GroupGradeLevelID, EmployeesHistoryList.IntegrationID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.LevelID, EmployeesHistoryList.PositionID, Payroll_" & lYearID & ".ConceptAmount, EmployeesGrades.GradePercentage From Payroll_" & lYearID & ", Employees, EmployeesChangesLKP, EmployeesHistoryList, EmployeesGrades, Levels, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lYearID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=EmployeesGrades.PayrollID) And (EmployeesHistoryList.EmployeeID=EmployeesGrades.EmployeeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Levels.StartDate<=EmployeesGrades.PayrollID) And (Levels.EndDate>=EmployeesGrades.PayrollID) And (Areas1.StartDate<=EmployeesGrades.PayrollID) And (Areas1.EndDate>=EmployeesGrades.PayrollID) And (Areas2.StartDate<=EmployeesGrades.PayrollID) And (Areas2.EndDate>=EmployeesGrades.PayrollID) And (EmployeesGrades.StartDate<=" & lYearID & "0101) And (EmployeesGrades.EndDate>=" & lYearID & "1231) And (Payroll_2011.RecordDate=EmployeesGrades.PayrollID) And (Payroll_2011.RecordID=EmployeesGrades.PayrollID) And (Payroll_" & lYearID & ".ConceptID=1) And (EmployeesGrades.Active=1) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFileName, ".xls", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			lErrorNumber = AppendTextToFile(Server.MapPath(sFileName), "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("No. Empleado,Empleado,Calificación,Plaza,Nivel Subnivel,Jornada,Zona,Sueldo base quincenal,Porcentaje,Sueldo base mensual,Sueldo base anual,Premio", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(Server.MapPath(sFileName), GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)

				asCellAlignments = Split(",CENTER,,CENTER,CENTER,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
					sTemp = " "
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
					Err.Clear
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeGrade").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("000000" & CStr(oRecordset.Fields("JobID").Value)), Len("000000"))) & """)"
					sTemp = ""
					sTemp = Right(("000" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000"))
					sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Left(sTemp, Len("N")) & "/" & Right(sTemp, Len("SN"))) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("GradePercentage").Value), 2, True, False, True) & "%"
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("ConceptAmount").Value) * 2), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("ConceptAmount").Value) * 24), 2, True, False, True)
					dAmount = FormatNumber((CDbl(oRecordset.Fields("ConceptAmount").Value) * 24 * CDbl(oRecordset.Fields("GradePercentage").Value) / 100), 2, True, False, False)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dAmount, 2, True, False, True)

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(Server.MapPath(sFileName), GetTableRowText(asRowContents, True, ""), sErrorDescription)
					'lErrorNumber = AppendTextToFile(Server.MapPath(Replace(sFileName, ".xls", ".txt")), (CStr(oRecordset.Fields("EmployeeID").Value) & ",32," & dAmount), sErrorDescription)
					lErrorNumber = AppendTextToFile(Server.MapPath(Replace(sFileName, ".xls", ".txt")), (CStr(oRecordset.Fields("EmployeeID").Value) & vbTab & dAmount), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			lErrorNumber = AppendTextToFile(Server.MapPath(sFileName), "</TABLE>", sErrorDescription)
			sTemp = "<B>Registrar los montos del reporte para la quincena de aplicación:</B><BR />"
			sTemp = sTemp & "<DIV>"' NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
				sTemp = sTemp & "<FORM NAME=""ConceptValuesFrm"" ID=""ConceptValuesFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
					sTemp = sTemp & "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" CLASS=""Lists"">"
						sTemp = sTemp & GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollID>=" & lYearID & "0000) And (PayrollID<=" & lYearID & "9999) And (PayrollTypeID=1) And (IsActive_3=1) And (IsClosed<>1)", "PayrollID Desc", "", "", sErrorDescription)
					sTemp = sTemp & "</SELECT><BR /><BR />"
					sTemp = sTemp & "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""ReasonIDHdn"" VALUE=""2"" />"
					sTemp = sTemp & "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & EMPLOYEES_EFFICIENCY_AWARD & """ />"
					sTemp = sTemp & "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeesMovements"" />"
					sTemp = sTemp & "<INPUT TYPE=""HIDDEN"" NAME=""FileName"" ID=""FileNameHdn"" VALUE=""" & Replace(sFileName, ".xls", ".txt") & """ />"
					sTemp = sTemp & "<INPUT TYPE=""SUBMIT"" NAME=""UploadFile"" ID=""UploadFileBtn"" VALUE=""Registrar Montos"" CLASS=""Buttons"" onClick=""this.form.action = 'UploadInfo.asp'"" />"
				sTemp = sTemp & "</FORM>"
			sTemp = sTemp & "</DIV>"
			Call DisplayInstructionsMessage("REGISTRO DE MONTOS", sTemp)

			lErrorNumber = ZipFolder(Server.MapPath(sFileName), Server.MapPath(Replace(sFileName, ".xls", ".zip")), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(Server.MapPath(sFileName), sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFileName, ".xls", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1211 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1221(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de carga de terceros institucionales
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1221"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sCondition
	Dim sQuery
	Dim lType

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "XXX", "Registration")

	oStartDate = Now()
	sQuery = "Select Employees.EmployeeID, Employees.EmployeeNumber," & _
			 " Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeName," & _
			 " CreditID, ContractNumber, AccountNumber, PaymentsNumber, CreditTypeShortName," & _
			 " Credits.StartDate, Credits.EndDate, Credits.UploadedRecordType," & _
			 " CreditTypeName, PaymentAmount, UploadedFileName, Comments" & _
			 " From Employees, Credits, CreditTypes" & _
			 " Where Employees.EmployeeID = Credits.EmployeeID" & _
			 " And Credits.CreditTypeID = CreditTypes.CreditTypeID" & sCondition

	sErrorDescription = "No se pudieron obtener los registros cargados del archivo de terceros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "CargaDeTerceros_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1221.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre archivo</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de crédito</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de movimiento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. de Pagos</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Comentarios</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UploadedFileName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("CreditTypeShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Indefinido</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
						lType = CLng(oRecordset.Fields("UploadedRecordType").Value)
						Select Case lType
							Case 1
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Alta</FONT></TD>"
							Case 2
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cambio</FONT></TD>"
							Case 3
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Baja</FONT></TD>"
						End Select
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PaymentsNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PaymentAmount").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("Comments").Value) & "</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1221 = lErrorNumber
	Err.Clear
End Function

Function BuildReports1221(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de Cargas de archivos de terceros
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReports1221"
	Dim sCondition
	Dim oRecordset
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los empleados registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From Employees, Jobs, Areas Where (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Employees.EmployeeID>-1) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

				Do While Not oRecordset.EOF
					'lErrorNumber = BuildReport1221(oRequest, oADODBConnection, CLng(oRecordset.Fields("EmployeeID").Value), True, sFilePath, sErrorDescription)
					lErrorNumber = BuildReport1221(oRequest, oADODBConnection, sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudo guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	BuildReports1221 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1222(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de rechazos en carga de terceros institucionales
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReports1222"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sCondition
	Dim sQuery
	Dim lType

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "XXX", "Registration")

	oStartDate = Now()
	sQuery = "Select EmployeeID, CreditTypeShortName, CreditTypeName, UploadedFileName, UploadedRecordType, UploadedRejectType, UploadedRecordLine, Comments" & _
			 " From CreditTypes, UploadThirdCreditsRejected" & _
			 " Where (UploadThirdCreditsRejected.CreditTypeID=CreditTypes.CreditTypeID)" & sCondition
	sQuery = sQuery & " Order By UploadedRejectType"

	sErrorDescription = "No se pudieron obtener los registros cargados del archivo de terceros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "CargaDeTerceros_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1221.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre archivo</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de crédito</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripcion del crédito</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de movimiento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de rechazo</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Linea</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Comentarios</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UploadedFileName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("CreditTypeShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("CreditTypeName").Value) & "</FONT></TD>"
						lType = CLng(oRecordset.Fields("UploadedRecordType").Value)
						Select Case lType
							Case 1
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Alta</FONT></TD>"
							Case 2
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cambio</FONT></TD>"
							Case 3
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Baja</FONT></TD>"
						End Select
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UploadedRejectType").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UploadedRecordLine").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("Comments").Value) & "</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1222 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1223(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de beneficiarios de pensión alimenticias
'         registrados para los empleados
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1223"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sCondition
	Dim sQuery
	Dim iNumEmp
	Dim iNumEmpAnt

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "XXX", "Registration")

	oStartDate = Now()
	sQuery = "Select Employees.EmployeeID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, ConceptAmount, ConceptQttyID, QttyName, AlimonyTypeName, EmployeesBeneficiariesLKP.StartDate, EmployeesBeneficiariesLKP.EndDate From Employees, EmployeesBeneficiariesLKP, QttyValues, AlimonyTypes" & _
			 " Where (Employees.EmployeeID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (AlimonyTypes.ConceptQttyID=QttyValues.QttyID) " & sCondition

	sErrorDescription = "No se pudieron obtener los registros de beneficiarios para el empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "BeneficiariosPension_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1223.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido materno del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Numero beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno del beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido materno del beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de pensión alimenticia</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad de descuento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Unidad</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de termino</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					iNumEmp = CLng(oRecordset.Fields("EmployeeNumber").Value)
					sRowContents = "<TR>"
						If iNumEmp <> iNumEmpAnt Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & "" & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">&nbps;</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryLastName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryLastName2").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					iNumEmpAnt = iNumEmp
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudo guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1223 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1224(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de empleados con pensión alimenticias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1224"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sCondition
	Dim sCondition2
	Dim sQuery
	Dim iNumEmp
	Dim iNumEmpAnt
	Dim sConceptNames

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "BeneficiaryStart", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "BeneficiaryEnd", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("BeneficiaryStart", "BeneficiaryEnd", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesBeneficiariesLKP.Start") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesBeneficiariesLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesBeneficiariesLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesBeneficiariesLKP.Start", 1, 1, vbBinaryCompare) & "))"
	oStartDate = Now()

	sQuery = "Select Employees.EmployeeID, EmployeeNumber, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName, ConceptAmount," & _
			 " BeneficiaryName + ' ' + BeneficiaryLastName + ' ' + BeneficiaryLastName2 As BeneficiaryFullName, BeneficiaryNumber, AlimonyTypeName, QttyName, AppliesToID," & _
			 " BeneficiaryID, EmployeesBeneficiariesLKP.StartDate, EmployeesBeneficiariesLKP.EndDate" & _
			 " From Employees, EmployeesBeneficiariesLKP, AlimonyTypes, QttyValues " & _
			 " Where Employees.EmployeeID = EmployeesBeneficiariesLKP.EmployeeID" & _
			 " and AlimonyTypes.AlimonyTypeID = EmployeesBeneficiariesLKP.AlimonyTypeID" & _
			 " and AlimonyTypes.ConceptQttyID = QttyValues.QttyID " & sCondition & sCondition2
	sErrorDescription = "No se pudieron obtener los registros de pensiones de empleados."

	sErrorDescription = "No se pudieron obtener los registros de beneficiarios para el empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "BeneficiariosPension_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1224.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Secuencia</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Numero beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del beneficiario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de pensión alimenticia</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad de descuento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Unidad</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Aplica a</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de termino</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sConceptNames = ""
					iNumEmp = CLng(oRecordset.Fields("EmployeeNumber").Value)
					sRowContents = "<TR>"
						If iNumEmp <> iNumEmpAnt Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & "" & "</FONT></TD>"
						End If
						Call GetConceptNamesFromAppliesToID(oADODBConnection, CStr(oRecordset.Fields("AppliesToID").Value), sConceptNames, sErrorDescription)
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryID").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryFullName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD WIDTH=1500 ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(sConceptNames)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					iNumEmpAnt = iNumEmp
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudo guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1224 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1225(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de cuentas bancarias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1225"
	Dim sHeaderContents
	Dim oRecordset
	Dim oRecordset1
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sCondition
	Dim sCurrentID
	Dim sQuery
	Dim lCount
	Dim lConceptCount
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollYear
	Dim sMaritalStatus
	Dim sGenderID
	Dim siConceptSIAmount
	Dim lAmount

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	lPayrollYear = Left(CStr(oRequest("YearID").Item), Len("0000"))
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	sCondition = sCondition & " And (Concepts.ConceptID = 120 Or Concepts.ConceptID = 122)"

	oStartDate = Now()

	sQuery = "Select EmployeesHistoryList.CompanyID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.PaymentCenterID, Employees.EmployeeID," & _
			" EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Employees.StartDate, CompanyShortName, CompanyName," & _
			" Employees.BirthDate, Employees.StartDate, CURP, SocialSecurityNumber, RFC, MaritalStatusID, GenderID," & _
			" Zones.ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2," & _
			" EmployeesHistoryList.JobID As JobNumber, PositionShortName, LevelShortName, GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID," & _
			" Concepts.ConceptID, ConceptShortName, IsDeduction, RecordDate, Payroll_" & lPayrollID & ".ConceptAmount, Zones.ZoneCode, PaymentCenters.EconomicZoneID" & _
			" From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies," & _
			" Areas, Positions, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, ZoneTypes" & _
			" Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID)" & _
			" And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID)" & _
			" And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate)" & _
			" And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID)" & _
			" And (AreasZones.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID)" & _
			" And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID)" & _
			" And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ")" & _
			" And (EmployeesHistoryList.EmployeeDate<=" & lForPayrollID & ") And (EmployeesHistoryList.EndDate>=" & lForPayrollID & ")" & _
			" And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ")" & _
			" And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ")" & _
			" And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ")" & _
			" And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & _
			" Order By Employees.EmployeeID, ConceptID"

	sErrorDescription = "No se pudieron obtener los registros de los empleados."

	Call DisplayTimeStamp("START: CONSULTA. " & sQuery)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1200Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sCurrentID = ""
			lConceptCount = 0
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sDocumentName = sFilePath & "Repcsi_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"

				Do While Not oRecordset.EOF
					If (StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0) Then
						lConceptCount = lConceptCount + CLng(oRecordset.Fields("ConceptAmount").Value)
						Select Case CInt(oRecordset.Fields("SocialSecurityNumber").Value) ' Campo35
							Case 0
								sMaritalStatus = "S"
							Case 1
								sMaritalStatus = "C"
							Case 2
								sMaritalStatus = "D"
							Case 3
								sMaritalStatus = "E"
							Case 4
								sMaritalStatus = "V"
							Case Else
								sMaritalStatus = "O"
						End Select
						Select Case CInt(oRecordset.Fields("GenderID").Value) ' Campo32
							Case 0
								sGenderID = "F"
							Case 1
								sGenderID = "M"
						End Select
						sRowContents = SizeText(CStr(oRecordset.Fields("CURP").Value), " ", 22, 1)										' Campo1
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("SocialSecurityNumber").Value), " ", 13, 1) 		' Campo2
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 14, 1)						' Campo3
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value), " ", 20, 1)			' Campo4
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeLastName2").Value), " ", 20, 1)		' Campo5
						Else
							sRowContents = sRowContents & "                   "															' Campo5
						End If
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeName").Value), " ", 20, 1)				' Campo6
						sRowContents = sRowContents & "                                        " ' Domicilio							' Campo7
						sRowContents = sRowContents & "                                        " ' Colonia								' Campo8
						sRowContents = sRowContents & "                                        " ' Poblacion							' Campo9
						sRowContents = sRowContents & "     " ' Codigo Postal															' Campo10 00000
						sRowContents = sRowContents & Left(CStr(oRecordset.Fields("PaymentCenterShortName").Value), Len("00"))			' Campo11 XX
						sRowContents = sRowContents & "               "																	' Campo12
						sRowContents = sRowContents & SizeText(CStr(FormatNumber(siConceptSIAmount, 2, True, False, True)), "0", 13, 0)	' Campo13 0000000.00
						sRowContents = sRowContents & SizeText(CStr(FormatNumber(CSng(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)), "0", 13, 0)	' Campo14 0000000.00
						sRowContents = sRowContents & "0000000.00"																		' Campo15 0000000.00
						sRowContents = sRowContents & "0000000.00"																		' Campo16 0000000.00
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value), " ", 6, -1)	' Campo17
						sRowContents = sRowContents & CStr(oRecordset.Fields("EconomicZoneID").Value)									' Campo18
						sRowContents = sRowContents & "1"																				' Campo19
						sRowContents = sRowContents & "A"																				' Campo20
						sRowContents = sRowContents & "P"																				' Campo21
						sRowContents = sRowContents & "A"																				' Campo22
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("StartDate").Value), " ", 8, 1)					' Campo23 99999999
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("StartDate2").Value), " ", 8, 1)					' Campo24 99999999
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("BirthDate").Value), " ", 8, 1)					' Campo25 99999999
						sRowContents = sRowContents & "0007"																			' Campo26 0000
						sRowContents = sRowContents & "999"																				' Campo27 000
						sRowContents = sRowContents & "082"																				' Campo28 000
						sRowContents = sRowContents & SizeText(CStr(FormatNumber(lConceptCount, 2, True, False, True)), "0", 13, 1)		' Campo29 0000000000.00
						sRowContents = sRowContents & "000"																				' Campo30 fill
						sRowContents = sRowContents & "0000"																			' Campo31 fill
						sRowContents = sRowContents & sGenderID																			' Campo32
						sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeName").Value), " ", 2, 1)				' Campo33 00
						sRowContents = sRowContents & "99"																				' Campo34
						sRowContents = sRowContents & SizeText(sMaritalStatus, " ", 1, 1)												' Campo35
						sRowContents = sRowContents & Left(CStr(oRecordset.Fields("AreaCode").Value), Len("00"))						' Campo36 00
						'sRowContents = sRowContents & Left(CStr(oRecordset.Fields("ConceptPercent").Value), Len("00"))					' Campo37 000.00
						sRowContents = sRowContents & Left(lAmount, Len("00"))															' Campo37 000.00
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lConceptCount = 0
						sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)' & CStr(oRecordset.Fields("RecordDate").Value)
						If Err.number <> 0 Then Exit Do
					End If
					'lConceptCount = lConceptCount + CLng(oRecordset.Fields("ConceptAmount").Value)
					If InStr(1, CStr(oRecordset.Fields("ConceptShortName").Value), "SI") > 0 Then
						Call GetConceptAmount(oADODBConnection, CInt(oRecordset.Fields("EmployeeID").Value), CInt(oRecordset.Fields("ConceptID").Value), lPayrollID, lAmount, sErrorDescription)
						siConceptSIAmount = CSng(oRecordset.Fields("ConceptAmount").Value)
					End If
					oRecordset.MoveNext
				Loop
			End If
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			oZonesRecordset.Close
		End If
	Else
		sErrorDescription = "Error al obtener las nominas para el mes seleccionado."
	End If
		'End If
	'End If

	Set oRecordset = Nothing
	BuildReport1225 = lErrorNumber
	Err.Clear
End Function
%>