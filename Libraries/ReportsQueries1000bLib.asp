<%
Function BuildReport1020(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditorías de nómina. Comparativo de nómina. Funcionarios con líquido mayor a la suma de sueldo mas compensación de la Qna. Reporte basado en la hoja 001179
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1020"
	Dim sHeaderContents
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim lAreaCurrentID
	Dim lEmployeeCurrentID
	Dim lCounter
	Dim lGlobalCounter
	Dim dTotal00
	Dim dTotal0103
	Dim sTemp
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Areas.", "Areas2."), "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "-17", "17")

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1020.htm"), sErrorDescription)
	If (Len(sHeaderContents) > 0) Then
		sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
		sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
		Response.Write sHeaderContents

		sErrorDescription = "No se pudieron obtener los conceptos de pago."
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.EmployeeTypeID=1) And (ConceptID In (0,1,3)) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID Order By Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeNumber, ConceptID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.EmployeeTypeID=1) And (ConceptID In (0,1,3)) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID Order By Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeNumber, ConceptID -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeTypeID=1) And (ConceptID In (0,1,3)) " & sCondition & " Group By Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID Order By Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeNumber, ConceptID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeTypeID=1) And (ConceptID In (0,1,3)) " & sCondition & " Group By Areas1.AreaID, Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, ConceptID Order By Areas1.AreaShortName, Areas1.AreaName, Employees.EmployeeNumber, ConceptID -->" & vbNewLine
		End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
					asCellAlignments = Split(",,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
					asColumnsTitles = Split("ÁREA,NOMBRE DE LA DELEGACIÓN,SUELDO BASE + COMPENSACIÓN, LÍQUIDO", ",")
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					lAreaCurrentID = -2
					lEmployeeCurrentID = -2
					lCounter = 0
					lGlobalCounter = 0
					dTotal00 = 0
					dTotal0103 = 0
					Do While Not oRecordset.EOF
						If lEmployeeCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							If lEmployeeCurrentID > -2 Then
								If dTotal00 > dTotal0103 Then
									sRowContents = Replace(sRowContents, "<CONCEPT_00 />", FormatNumber(dTotal00, 2, True, False, True))
									sRowContents = Replace(sRowContents, "<CONCEPT_0103 />", FormatNumber(dTotal0103, 2, True, False, True))
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									lCounter = lCounter + 1
									lGlobalCounter = lGlobalCounter + 1
								End If
								dTotal00 = 0
								dTotal0103 = 0
							End If
							sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)

							Err.Clear
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_0103 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_00 />"
							lEmployeeCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						End If
						If lAreaCurrentID <> CLng(oRecordset.Fields("AreaID").Value) Then
							If lCounter > 0 Then
								asRowContents = Split(("<SPAN COLS=""3"" /><B>Registros</B>" & TABLE_SEPARATOR & "<B>" & lCounter & "</B>"), TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
							asRowContents = Split(("<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""3"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"), TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lCounter = 0
							lAreaCurrentID = CLng(oRecordset.Fields("AreaID").Value)
						End If
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 0
								dTotal00 = CDbl(oRecordset.Fields("TotalAmount").Value)
							Case Else
								dTotal0103 = dTotal0103 + CDbl(oRecordset.Fields("TotalAmount").Value)
						End Select
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close

					If dTotal00 > dTotal0103 Then
						sRowContents = Replace(sRowContents, "<CONCEPT_00 />", FormatNumber(dTotal00, 2, True, False, True))
						sRowContents = Replace(sRowContents, "<CONCEPT_0103 />", FormatNumber(dTotal0103, 2, True, False, True))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCounter = lCounter + 1
						lGlobalCounter = lGlobalCounter + 1
					End If

					If lCounter > 0 Then
						asRowContents = Split(("<SPAN COLS=""3"" /><B>Registros</B>" & TABLE_SEPARATOR & "<B>" & lCounter & "</B>"), TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If

					asRowContents = Split(("<SPAN COLS=""3"" /><B>Total global de empleados con liquido mayor a sueldo base + compensación:</B>" & TABLE_SEPARATOR & "<B>" & lGlobalCounter & "</B>"), TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				Response.Write "</TABLE>"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oRecordset.Close
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "La plantilla del reporte no existe."
	End If

	Set oRecordset = Nothing
	BuildReport1020 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1021(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Auditorías de nómina. Comparativo de nómina. Totales por nómina. Reporte basado en la hoja 001180
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1021"

	Dim sHeaderContents
	Dim iMonth
	Dim iYear
	Dim iIndex
	Dim asPayrollIDs
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sTitles
	Dim lErrorNumber
	Dim sActualRecord
	Dim sNextRecord
	Dim iTotalRecords
	Dim dPerception
	Dim dDeducction
	Dim dTotalPayment
	Dim iTotalRecords2
	Dim dPerception2
	Dim dDeducction2
	Dim dTotalPayment2
	Dim bNoRecords
	Dim iCount
	iCount = 0

	iMonth = CInt(Mid(oRequest.Item("PayrollID").Item,5,2))
	iYear = CInt(Mid(oRequest.Item("PayrollID").Item,1,4))
	If iMonth = 12 Then
		iMonth = 1
		iYear = iYear + 1
	Else
		iMonth = iMonth + 1
	End If

	sErrorDescription = "No se pudieron obtener el total pagado por percepciones para el periodo especificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollDate>=" & oRequest.Item("PayrollID").Item & ") And (PayrollDate<=" & (iYear & Right(("0" & iMonth), Len("00"))) & "99) And (PayrollTypeID=1) Order By PayrollID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrollIDs = ""
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			iCount = iCount + 1
			asPayrollIDs = asPayrollIDs & CStr(oRecordset.Fields("PayrollID").Value) & ";"
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	oRecordset.Close

	sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1021.htm"), sErrorDescription)
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

		bNoRecords = True
		sErrorDescription = "No se pudieron obtener los montos de pago."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(0) & " Where ConceptID = -1", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				iTotalRecords = CLng(oRecordset.Fields("Total").Value)
				dPerception = CDbl(oRecordset.Fields("TotalAmount").Value)
				bNoRecords = False
			End If
			oRecordset.Close
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(0) & " Where ConceptID = -2", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					dDeducction = CDbl(oRecordset.Fields("TotalAmount").Value)
					bNoRecords = False
				End If
				oRecordset.Close
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(0) & " Where ConceptID = 0", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						dTotalPayment = CDbl(oRecordset.Fields("TotalAmount").Value)
						bNoRecords = False					
						Response.Write "<TABLE BORDER="""
							If Not bForExport Then
								Response.Write "0"
							Else
								Response.Write "1"
							End If
						Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
						asCellAlignments = Split("CENTER,LEFT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
						sTitles = "FECHA<BR />PAGO,NOMBRE,REGISTROS,TOTAL<BR />DEVENGADOS,TOTAL<BR />RETENIDO,LIQUIDO"
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
						sRowContents = asPayrollIDs(0)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & asPayrollIDs(0)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(iTotalRecords, 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dPerception, 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dDeducction, 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dTotalPayment, 2, True, False, True)
						asCellAlignments = Split("RIGHT,RIGHT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
				End If
			End If
		End If

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(1) & " Where ConceptID = -1", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				iTotalRecords2 = CLng(oRecordset.Fields("Total").Value)
				dPerception2 = CDbl(oRecordset.Fields("TotalAmount").Value)
				bNoRecords = False
			End If
			oRecordset.Close
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(1) & " Where ConceptID = -2", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					dDeducction2 = CDbl(oRecordset.Fields("TotalAmount").Value)
					bNoRecords = False
				End If
				oRecordset.Close
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As Total, Sum(ConceptAmount) As TotalAmount From Payroll_" & asPayrollIDs(1) & " Where ConceptID = 0", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						dTotalPayment2 = CDbl(oRecordset.Fields("TotalAmount").Value)
						bNoRecords = False
						sRowContents = asPayrollIDs(1)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & asPayrollIDs(1)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(iTotalRecords2, 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dPerception2, 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dDeducction2, 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber(dTotalPayment2, 2, True, False, True)
						asCellAlignments = Split("RIGHT,RIGHT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						
						sRowContents = ""
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & "Diferencia entre quincenas"
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((iTotalRecords-iTotalRecords2), 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((dPerception-dPerception2), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((dDeducction-dDeducction2), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & FormatNumber((dTotalPayment-dTotalPayment2), 2, True, False, True)
						asCellAlignments = Split("RIGHT,RIGHT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						Response.Write "</TABLE>"
					End If
				End If
			End If
		End If
	End If

	If bNoRecords = True Then
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	BuildReport1021 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1022(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Empleados con líquidos mayores al monto introducido por el usuario. Reporte basado en la hoja 001185
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1022"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim oRecordset
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

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	If (InStr(1, sCondition, "Companies.", vbBinaryCompare) > 0) Then
		sTableNames = ", Companies"
	End If

    oStartDate = Now()
	sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeLastName, EmployeeLastName2, EmployeeName, JobNumber, PositionShortName, LevelShortName, Jobs.WorkingHours, EmployeeTypeShortName, PositionTypeShortName, GroupGradeLevelShortName, Areas.AreaCode, PaymentCenters.AreaCode As PaymentCenterShortName, ServiceShortName, (Percepciones.ConceptAmount - Deducciones.ConceptAmount) As TotalPayment From Employees, Jobs, Levels, EmployeeTypes, PositionTypes, GroupGradeLevels, Services, Positions, Areas As PaymentCenters, Areas, Payroll_" & lPayrollID & " As Percepciones, Payroll_" & lPayrollID & " As Deducciones" & sTableNames & " Where (Percepciones.ConceptID=-1) And (Deducciones.ConceptID=-2) And (Jobs.JobID=Employees.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Levels.LevelID=Employees.LevelID) And (EmployeeTypes.EmployeeTypeID=Employees.EmployeeTypeID) And (PositionTypes.PositionTypeID=Employees.PositionTypeID) And (GroupGradeLevels.GroupGradeLevelID=Employees.GroupGradeLevelID) And (Services.ServiceID=Employees.ServiceID) And (PaymentCenters.AreaID=Employees.PaymentCenterID) And (Percepciones.EmployeeID=Employees.EmployeeID) And (Deducciones.EmployeeID=Employees.EmployeeID) " & sCondition & " Order by Areas.AreaCode", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, EmployeeLastName, EmployeeLastName2, EmployeeName, JobNumber, PositionShortName, LevelShortName, Jobs.WorkingHours, EmployeeTypeShortName, PositionTypeShortName, GroupGradeLevelShortName, Areas.AreaCode, PaymentCenters.AreaCode As PaymentCenterShortName, ServiceShortName, (Percepciones.ConceptAmount - Deducciones.ConceptAmount) As TotalPayment From Employees, Jobs, Levels, EmployeeTypes, PositionTypes, GroupGradeLevels, Services, Positions, Areas As PaymentCenters, Areas, Payroll_" & lPayrollID & " As Percepciones, Payroll_" & lPayrollID & " As Deducciones" & sTableNames & " Where (Percepciones.ConceptID=-1) And (Deducciones.ConceptID=-2) And (Jobs.JobID=Employees.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Levels.LevelID=Employees.LevelID) And (EmployeeTypes.EmployeeTypeID=Employees.EmployeeTypeID) And (PositionTypes.PositionTypeID=Employees.PositionTypeID) And (GroupGradeLevels.GroupGradeLevelID=Employees.GroupGradeLevelID) And (Services.ServiceID=Employees.ServiceID) And (PaymentCenters.AreaID=Employees.PaymentCenterID) And (Percepciones.EmployeeID=Employees.EmployeeID) And (Deducciones.EmployeeID=Employees.EmployeeID) " & sCondition & " Order by Areas.AreaCode -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()		
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1022.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				End If

				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TR>"
				sRowContents = sRowContents & "<TD>EMP.</TD>"
				sRowContents = sRowContents & "<TD>NOMBRE</TD>"
				sRowContents = sRowContents & "<TD>PLAZA</TD>"
				sRowContents = sRowContents & "<TD>PUESTO</TD>"
				sRowContents = sRowContents & "<TD>N/SN</TD>"
				sRowContents = sRowContents & "<TD>JOR</TD>"
				sRowContents = sRowContents & "<TD>TAB</TD>"
				sRowContents = sRowContents & "<TD>TPTO</TD>"
				sRowContents = sRowContents & "<TD>GGN</TD>"
				sRowContents = sRowContents & "<TD>C_TRAB</TD>"
				sRowContents = sRowContents & "<TD>C_PAGO</TD>"
				sRowContents = sRowContents & "<TD>SERVICIO</TD>"
				sRowContents = sRowContents & "<TD>LIQUIDO</TD>"
				sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
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
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "<TD>"
					sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("TotalPayment").Value), 2, True, False, True)
					sRowContents = sRowContents & "</TD>"
					sRowContents = sRowContents & "</TR>"

					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				sRowContents = "</TABLE>"
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	BuildReport1022 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1023(oRequest, oADODBConnection, bShowTotals, bForExport, sErrorDescription)
'************************************************************
'Purpose: Diferencias de empleados por unidad administrativa. Reporte basado en la hoja 001172
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bShowTotals, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1023"
	

	Set oRecordset = Nothing
	BuildReport1023 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1024(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal con conceptos. Reporte basado en la hoja 001221 
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1024"
	Dim sCondition
	Dim sCompanyCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCurrentID
	Dim oRecordset
	Dim asCompanies
	Dim iIndex
	Dim asColumnsTitles
	Dim sRowContents
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sSourceFolderPath
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

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCompanyCondition = ""
	If Len(oRequest("CompanyID").Item) > 0 Then
		sCompanyCondition = " And (CompanyID In (" & oRequest("CompanyID").Item & "))"
	End If

	sDate = GetSerialNumberForDate("")
	lTotalEmployees = 0
	dTotalAmount = 0
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		sSourceFolderPath  = Server.MapPath(TEMPLATES_PATH & "Images")
		sSourceFolderPath = sSourceFolderPath & "\"

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
			oStartDate = Now()

			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyName From Companies Where (ParentID>=0) And (EndDate=30000000) And (Active=1) " & sCompanyCondition & " Order By CompanyShortName", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				asCompanies = ""
				Do While Not oRecordset.EOF
					asCompanies = asCompanies & CStr(oRecordset.Fields("CompanyID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("CompanyName").Value) & SECOND_LIST_SEPARATOR & "1" & SECOND_LIST_SEPARATOR & "FUNCIONARIOS" & LIST_SEPARATOR
					asCompanies = asCompanies & CStr(oRecordset.Fields("CompanyID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("CompanyName").Value) & SECOND_LIST_SEPARATOR & "0,2,3,4,5,6" & SECOND_LIST_SEPARATOR & "OPERATIVOS" & LIST_SEPARATOR
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				asCompanies = Left(asCompanies, (Len(asCompanies) - Len(LIST_SEPARATOR)))
				asCompanies = Split(asCompanies, LIST_SEPARATOR)
			End If
			
			For iIndex = 0 To UBound(asCompanies)
				asCompanies(iIndex) = Split(asCompanies(iIndex), SECOND_LIST_SEPARATOR)
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, PositionShortName, LevelShortName, ConceptShortName, IsDeduction, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Jobs, Positions, Levels, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (Jobs.PositionID=Positions.PositionID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryList.CompanyID=" & asCompanies(iIndex)(0) & ") And (EmployeesHistoryList.EmployeeTypeID In (" & asCompanies(iIndex)(2) & ")) " & Replace(sCondition, "Companies.", "EmployeesHistoryList.") & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, PositionShortName, LevelShortName, ConceptShortName, IsDeduction Order By EmployeesHistoryList.EmployeeNumber, ConceptShortName", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					lCurrentID = -2
					lTotalEmployees = 0
					dTotalAmount = 0
					If Not oRecordset.EOF Then
						bEmpty = False
						sRowContents = "<BR /><B>EMPRESA:" & CleanStringForHTML(asCompanies(iIndex)(1)) & "</B><BR /><BR />"
						sRowContents = sRowContents & "<B>" & CleanStringForHTML(asCompanies(iIndex)(3)) & "</B><BR />"
						sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>No. EMP.</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>NOMBRE</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>ADSCRIP.</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>RFC</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>PUESTO</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Niv/SubN</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>CONCEPTO</B></FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>MONTO</B></FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							asCellAlignments = Split(",,,,,CENTER,,RIGHT", ",", -1, vbBinaryCompare)
							Do While Not oRecordset.EOF
								If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
									lTotalEmployees = lTotalEmployees + 1
									lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								End If
								sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & """)"
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
								If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
									dTotalAmount = dTotalAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
								Else
									dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
								End If

								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								lErrorNumber = AppendTextToFile(sDocumentName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
							oRecordset.Close

							sRowContents = "<SPAN COLS=""8"" />TOTAL " & CleanStringForHTML(asCompanies(iIndex)(1)) & " " & CleanStringForHTML(asCompanies(iIndex)(3)) & ": " & lTotalEmployees
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sDocumentName, GetTableRowText(asRowContents, True, ""), sErrorDescription)

							sRowContents = "<SPAN COLS=""8"" />MONTO " & CleanStringForHTML(asCompanies(iIndex)(3)) & ": " & FormatNumber(dTotalAmount, 2, True, False, True)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sDocumentName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
						lErrorNumber = AppendTextToFile(sDocumentName, "</TABLE><BR />", sErrorDescription)
					End If
				End If
			Next

			If Not bEmpty Then
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
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		End If
	End If
	Set oRecordset = Nothing
	BuildReport1024 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1026(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de incidencias registradas para el personal filtrado por
'         número de empleado, áreas y período específico
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1026"
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

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asAbsenceNames
	Dim asPath
	Dim iCount
	Dim aiAbscenceTotals
	Dim aiAbscenceGrandTotals
	Dim iIndex
	Dim sAbsenceShortName
	Dim bFirst
	Dim lTotal
	Dim lTotalForArea
	Dim lTotalForReport
	Dim iMin
	Dim iMax

	sQuery = "Select EmployeesAbsencesLKP.EmployeeID, Employees.EmployeeNumber, Employees.PaymentCenterID, EmployeesAbsencesLKP.AbsenceID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," & _
			 " EmployeesAbsencesLKP.AppliedDate, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, EmployeesAbsencesLKP.RegistrationDate, VacationPeriod, EmployeesAbsencesLKP.DocumentNumber, EmployeesAbsencesLKP.AbsenceHours," & _
			 " EmployeesAbsencesLKP.JustificationID, J.JustificationShortName, EmployeesAbsencesLKP.Reasons, EmployeesAbsencesLKP.Removed, EmployeesAbsencesLKP.JustificationID As AbsenceJustified, EmployeesAbsencesLKP.Active, Employees.JourneyID," & _
			 " EmployeesAbsencesLKP.AppliedRemoveDate, A.AbsenceShortName, A.AbsenceName, A.IsJustified, A.JustificationID As WithJustification, Users.UserLastName + ' ' + Users.UserName As UserFullName," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName" & _
			 " From Employees, EmployeesAbsencesLKP, Absences As A, Justifications As J, Users, Areas, Areas As PaymentCenters, Jobs," & _
			 " Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.JustificationID=J.JustificationID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID=A.AbsenceID) And (EmployeesAbsencesLKP.AddUserID=Users.UserID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID < 100)"

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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(AbsenceID) As Max From Absences", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value)
		End If
	End If
	For iMin = 0 To iMax
		asAbsenceNames = asAbsenceNames & LIST_SEPARATOR & ""
		aiAbscenceTotals = aiAbscenceTotals & LIST_SEPARATOR & "0"
		aiAbscenceGrandTotals = aiAbscenceGrandTotals & LIST_SEPARATOR & "0"
	Next
	asAbsenceNames = Split(asAbsenceNames, LIST_SEPARATOR)
	aiAbscenceTotals = Split(aiAbscenceTotals, LIST_SEPARATOR)
	aiAbscenceGrandTotals = Split(aiAbscenceGrandTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiAbscenceTotals(iIndex) = 0
		aiAbscenceGrandTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, AbsenceShortName From Absences Where (AbsenceID>-1) And (AbsenceID<100) Order By AbsenceID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asAbsenceNames(CInt(oRecordset.Fields("AbsenceID").Value)) = SizeText(CStr(CleanStringForHTML(oRecordset.Fields("AbsenceShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "OcurredDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("OcurredDate", "EndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.Ocurred") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesAbsencesLKP.Ocurred", 1, 1, vbBinaryCompare) & "))"
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
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
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1026.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
					lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				End If
				iCount = 0
				lCurrentPaymentCenterID = -2
				lTotalForReport = 0
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
							sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
							sRowContents = sRowContents & "<TD>TOTAL</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							lTotalForArea = 0
							For iIndex = 0 To UBound(aiAbscenceTotals)
								lTotal = CInt(aiAbscenceTotals(iIndex))
								If lTotal > 0 Then
									lTotalForArea = lTotalForArea + lTotal
									lTotalForReport = lTotalForReport + lTotal
									sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
										sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
									sRowContents = sRowContents & "</FONT></TR>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							Next
							For iIndex = 0 To UBound(aiAbscenceTotals)
								aiAbscenceTotals(iIndex) = 0
							Next
							sRowContents = "</TABLE>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							sRowContents = "<BR /><B>REGISTROS TOTALES POR CENTRO DE TRABAJO: " & lTotalForArea & "</B><BR />"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
							sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
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
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de ocurrencia</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de término</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Periodo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave del centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Estatus</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Justificación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Q. de aplicación de la justificación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. de documento</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del usuario</FONT></TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CInt(oRecordset.Fields("AbsenceID").Value)
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeFullName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("=T(""" & CStr(oRecordset.Fields("AbsenceShortName").Value)) & """)</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "</FONT></TD>"
						If (Not VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) Or ((CInt(oRecordset.Fields("JourneyID").Value)=21) Or (CInt(oRecordset.Fields("JourneyID").Value)=22) Or (CInt(oRecordset.Fields("JourneyID").Value)=23))) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>"
						Else
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("A la fecha") & "</FONT></TD>"
							Else
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
							End If
						End If
						If CInt(oRecordset.Fields("VacationPeriod").Value) <= 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("NA") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Left(CStr(oRecordset.Fields("VacationPeriod").Value), Len("0000")) & "-" & Right(CStr(oRecordset.Fields("VacationPeriod").Value), Len("0"))) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("Aplicada") & "</FONT></TD>"
						ElseIf CInt(oRecordset.Fields("Active").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("En proceso") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("Cancelada") & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("JustificationShortName").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("JustificationID").Value) = -1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("NA") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedRemoveDate").Value)) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value)) & "</FONT></TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					If CInt(oRecordset.Fields("JustificationID").Value) <> -1 Then
						aiAbscenceTotals(CInt(oRecordset.Fields("JustificationID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("JustificationID").Value)) + 1
						aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) + 1
					Else
						aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
						aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
					End If
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
					sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
					sRowContents = sRowContents & "<TD>TOTAL</TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lTotalForArea = 0
					For iIndex = 0 To UBound(aiAbscenceTotals)
						lTotal = CInt(aiAbscenceTotals(iIndex))
						If lTotal > 0 Then
							lTotalForArea = lTotalForArea + lTotal
							lTotalForReport = lTotalForReport + lTotal
							sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
					For iIndex = 0 To UBound(aiAbscenceTotals)
						aiAbscenceTotals(iIndex) = 0
					Next
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					sRowContents = "<BR /><B>REGISTROS TOTALES POR CENTRO DE TRABAJO: " & lTotalForArea & "</B><BR />"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				End If
				sRowContents = "<BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>TOTALES DEL REPORTE</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
				sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
				sRowContents = sRowContents & "<TD>TOTAL</TD>"
				sRowContents = sRowContents & "</FONT></TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				For iIndex = 0 To UBound(aiAbscenceGrandTotals)
					lTotal = CInt(aiAbscenceGrandTotals(iIndex))
					If lTotal > 0 Then
						sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
							sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				Next
				For iIndex = 0 To UBound(aiAbscenceGrandTotals)
					aiAbscenceGrandTotals(iIndex) = 0
				Next
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>REGISTROS TOTALES DEL REPORTE: " & lTotalForReport & "</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.Close
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1026 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1027(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Listado para impresión de cheques.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1027"
	Dim sGeneralHeader
	Dim sEmployeeHeader
	Dim sContents
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lStartPayrollDate
	Dim lCurrentEmployeeID
	Dim asStateNames
	Dim asPath
	Dim sPeriod
	Dim asConceptsP
	Dim asConceptsD
	Dim iIndex
	Dim jIndex
	Dim sRowContents
	Dim sConcepts
	Dim bPayrollIsClosed
	Dim oRecordset
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber
	Dim sQuery
	Dim lPerceptions
	Dim lDeductions
	Dim lTotal
	Dim bFirstEmployee
	Dim yPositionForConcepts

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
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

		Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
		lStartPayrollDate = GetPayrollStartDate(lForPayrollID)
		If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'			sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
			sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		End If

		Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

		If bPayrollIsClosed Then
			sQuery = "Select EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.PaymentCenterID, Employees.EmployeeID," & _
					 " EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2," & _
					 " RFC, CURP, SocialSecurityNumber, Employees.StartDate," & _
					 " CompanyShortName, CompanyName, ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName," & _
					 " PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, EmployeesHistoryListForPayroll.JobID As JobNumber," & _
					 " PositionShortName, PositionName, LevelShortName, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, " & _
					 " RecordDate, ConceptAmount, CheckNumber, EmployeesHistoryListForPayroll.BankID From Payments, Payroll_" & lPayrollID & ", Concepts," & _
					 " Employees, EmployeesHistoryListForPayroll, Companies, Areas, Positions, Levels," & _
					 " Areas As PaymentCenters, Zones, ZoneTypes Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID)" & _
					 " And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID)" & _
					 " And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID)" & _
					 " And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID)" & _
					 " And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ")" & _
					 " And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID)" & _
					 " And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID)" & _
					 " And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
					 " And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ")" & _
					 " And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ")" & _
					 " And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ")" & _
					 " And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ")" & _
					 " And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ")" & _
					 " And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ")" & _
					 " And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ")" & _
					 " And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ")" & _
					 " " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By Payments.CheckNumber, OrderInList, RecordDate, RecordID"
		Else
			sQuery = "Select EmployeesHistoryList.CompanyID, EmployeesHistoryList.PaymentCenterID, Employees.EmployeeID," & _
					 " EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2," & _
					 " RFC, CURP, SocialSecurityNumber, Employees.StartDate," & _
					 " CompanyShortName, CompanyName, ZonePath, PaymentCenters.AreaCode As PaymentCenterShortName," & _
					 " PaymentCenters.AreaName As PaymentCenterName, ZoneTypeID2, EmployeesHistoryList.JobID As JobNumber," & _
					 " PositionShortName, PositionName, LevelShortName, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, " & _
					 " RecordDate, ConceptAmount, CheckNumber, BankID From Payments, BankAccounts, Payroll_" & lPayrollID & ", Concepts," & _
					 " Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels," & _
					 " Areas AS PaymentCenters, Zones, ZoneTypes Where (Payments.AccountID=BankAccounts.AccountID)" & _
					 " And (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID)" & _
					 " And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID)" & _
					 " And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID)" & _
					 " And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID)" & _
					 " And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID)" & _
					 " And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate)" & _
					 " And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate)" & _
					 " And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID)" & _
					 " And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID)" & _
					 " And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
					 " And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Payments.PaymentDate=" & lPayrollID & ")" & _
					 " And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate)" & _
					 " And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ")" & _
					 " And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ")" & _
					 " And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ")" & _
					 " And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ")" & _
					 " And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ")" & _
					 " And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ")" & _
					 " And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ")" & _
					 " And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & _
					 " Order By Payments.CheckNumber, OrderInList, RecordDate, RecordID"
		End If
		sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bFirstEmployee = True
				sDate = GetSerialNumberForDate("")
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
				If lErrorNumber = 0 Then
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()
						
					lCurrentCompanyID = -1
					lCurrentPaymentCenterID = -1
					lCurrentEmployeeID = -1
					sContents = ""
					sConcepts = ""
					lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\froman Times New Roman;}}\fs16", sErrorDescription)
						Do While Not oRecordset.EOF
							If lCurrentEmployeeID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								yPositionForConcepts=2500
								If Not bFirstEmployee Then
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\sbkpage{\*\atnid S A L T O  D E  S E C C I O N}", sErrorDescription)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\sect\sectd{\*\atnid N U E V A  S E C C I O N}", sErrorDescription)
								End If
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy397 \absw1247{\*\atnid NO.EMP.}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx1701 \posy397 \absw8392{\*\atnid NOMBRE}", sErrorDescription)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								Else
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								End If
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy879 \absw2240{\*\atnid RFC}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx2665 \posy879 \absw5075{\*\atnid CURP}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx7740 \posy879 \absw3856{\*\atnid NSS}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy1361 \absw2240{\*\atnid CLAVE_PUESTO}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 7, 1)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx2665 \posy1361 \absw3515{\*\atnid DESC_PUESTO}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 60, 1)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx6209 \posy1361 \absw737{\*\atnid NIV_SUB}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), Left(Right(("00" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000")), Len("00")), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx6917 \posy1361 \absw567{\*\atnid RG_MX}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "2/0", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx8051 \posy1361 \absw567{\*\atnid CG_RP}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx7484 \posy1361 \absw1814{\*\atnid CLAVE_PRESUPUESTAL}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "4588", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx9866 \posy1361 \absw1701{\*\atnid CLAVE_DISTRIBUCION}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "4588", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy1814 \absw1956{\*\atnid FECHA_INGRESO}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx2296 \posy1814 \absw1956{\*\atnid FECHA_NOMINA}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), DisplayNumericDateFromSerialNumber(lForPayrollID), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx4224 \posy1814 \absw3742{\*\atnid PERIODO_PAGO}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), DisplayNumericDateFromSerialNumber(lStartPayrollDate) & " AL " & DisplayNumericDateFromSerialNumber(lForPayrollID), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx6350 \posy12446 \absw5613{\*\atnid FECHA_CHEQUE}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), DisplayNumericDateFromSerialNumber(lForPayrollID), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx2212 \posy13100 \absw8500{\*\atnid NOMBRE_CHEQUE}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\fs18 \b", sErrorDescription)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								Else
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								End If
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx9660 \posy12820 \absw1985{\*\atnid IMPORTE_CHEQUE}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), FormatNumber(lTotal, 2, True, False, True), sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx6379 \posy14742 \absw2778{\*\atnid FIRMA1}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "C:\'5c\'5cSIAP\'5c\'5cTemplates\'5c\'5cImages\'5c\'5c0202.jpg", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx8959 \posy14742 \absw2778{\*\atnid FIRMA2}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "C:\'5c\'5cSIAP\'5c\'5cTemplates\'5c\'5cImages\'5c\'5c0202.jpg", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								lCurrentEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
								bFirstEmployee = False
							End If

							Select Case CLng(oRecordset.Fields("ConceptID").Value)
								Case -2
									lDeductions = CDbl(oRecordset.Fields("ConceptAmount").Value)
								Case -1
									lPerceptions = Replace(sContents, "<PERCEPTIONS />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx7966 \posy1814 \absw1814{\*\atnid BASE_GRAVABLE}", sErrorDescription)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), FormatNumber(lPerceptions, 2, True, False, True), sErrorDescription)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								Case 0
									sContents = Replace(sContents, "<TOTAL />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx9781 \posy1814 \absw1701{\*\atnid NETO_A_PAGAR}", sErrorDescription)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), FormatNumber(lTotal, 2, True, False, True), sErrorDescription)
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
								Case Else
									If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
										asConceptsP = asConceptsP & "P " & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptShortName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy" & yPositionForConcepts & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CStr(oRecordset.Fields("ConceptShortName").Value), sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx1100 \posy" & yPositionForConcepts & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CStr(oRecordset.Fields("ConceptName").Value), sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx4824 \posy" & yPositionForConcepts & " \absw1814{\*\atnid IMPORTE}", sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), CStr(oRecordset.Fields("ConceptAmount").Value), sErrorDescription)
										lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
										yPositionForConcepts = yPositionForConcepts + 250
									Else
										asConceptsD = asConceptsD & "D " & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptShortName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & SECOND_LIST_SEPARATOR & lTempStartDate & "." & lTempEndDate & LIST_SEPARATOR
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx340 \posy2500 \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "45", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx1100 \posy2500 \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "Pensión Alimenticia", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "{\pard \pvpg\phpg \posx4824 \posy2500 \absw1814{\*\atnid IMPORTE}", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "12345.67", sErrorDescription)
										'lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "\par}", sErrorDescription)
									End If
							End Select
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
					lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".rtf"), "}", sErrorDescription)

					lErrorNumber = ZipFile(Server.MapPath(sFileName & ".rtf"), Server.MapPath(sFileName & ".zip"), sErrorDescription)
					If lErrorNumber = 0 Then
						Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
						sErrorDescription = "No se pudieron guardar la información del reporte."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						lErrorNumber = DeleteFile(Server.MapPath(sFileName & ".rtf"), sErrorDescription)
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
	End If

	Set oRecordset = Nothing
	BuildReport1027 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1028(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de cifras del bimestre)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1028"
	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sPayrollDate
	Dim sQuery
	Dim sRowContents
	Dim sTruncate
	Dim lCpto_01
	Dim lCpto_04
	Dim lCpto_05
	Dim lCpto_06
	Dim lCpto_07
	Dim lCpto_08
	Dim lCpto_11
	Dim lCpto_44
	Dim lCpto_B2
	Dim lCpto_7S
	Dim lLastCompany
	Dim lDeductions
	Dim lErrorNumber
	Dim lForPayrollID
	Dim lPayrollID
	Dim lPerceptions
	Dim lPeriodID
	Dim lTotal
	Dim oRecordset
	Dim oRecordsetSumary
	Dim alPayrolls
	Dim asConcepts
	Dim asConcepts1
	Dim asConcepts2
	Dim asCompanies
	Dim iIndex
	Dim jIndex

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	'Extracción de catálogo de conceptos
	sQuery = "Select ConceptID, ConceptShortName From Concepts Where ConceptShortName in ('01','04','05','06','07','08','11','44','7S','B2')"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asConcepts = oRecordset.getRows()
	Set oRecordset = Nothing
	'Extracción de catállogo de empresas
	sQuery = "Select CompanyID, CompanyShortName, CompanyName From Companies where CompanyID > 0"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asCompanies = oRecordset.getRows()
	Set oRecordset = Nothing

	sErrorDescription = "No se pudo obtener la información para generar el reporte de cifras del bimestre"
	sQuery = "Select * From DM_Hist_Binmar Order By societyid, CompanyID, PaymentDate"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".htm"
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1028_1.htm"), sErrorDescription)
			sHeaderContents = Replace(sHeaderContents, "<PERIOD_ID />", lPeriodID)
			lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
			sRowContents = sRowContents & "<BR />"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""25%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""25%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>" & CStr(oRecordset.Fields("CompanyID").Value) & "</TD><TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>" & asCompanies(2,CLng(oRecordset.Fields("CompanyID").Value)-1) & "</B></FONT></TD>"
				sRowContents = sRowContents & "</TR>"
			sRowContents = sRowContents & "</TABLE>"
			sRowContents = sRowContents & "<BR />"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			lLastCompany = oRecordset.Fields("CompanyID").Value
			Do While Not oRecordset.EOF
				If (CLng(oRecordset.Fields("CompanyID").Value) <> CLng(lLastCompany)) Then
					sQuery = "Select Sum(Income) As Income, Sum(Deductions) As Deductions, Sum(NetIncome) As NetIncome, Sum(Cpt_n01) As Cpt_n01, Sum(Cpt_a01) As Cpt_a01, Sum(Cpt_n04) As Cpt_n04, Sum(Cpt_a04) As Cpt_a04, Sum(Cpt_n05) As Cpt_n05, Sum(Cpt_a05) As Cpt_a05, Sum(Cpt_n06) As Cpt_n06, Sum(Cpt_a06) As Cpt_a06, Sum(Cpt_n07) As Cpt_n07, Sum(Cpt_a07) As Cpt_a07, Sum(Cpt_n08) As Cpt_n08, Sum(Cpt_a08) As Cpt_a08, Sum(Cpt_n11) As Cpt_n11, Sum(Cpt_a11) As Cpt_a11, Sum(Cpt_n44) As Cpt_n44, Sum(Cpt_a44) As Cpt_a44, Sum(Cpt_n7s) As Cpt_n7s, Sum(Cpt_a7s) As Cpt_a7s, Sum(Cpt_nb2) As Cpt_nb2, Sum(Cpt_ab2) As Cpt_ab2 From DM_Hist_Binmar Where (CompanyID=" & lLastCompany & ")"
					sErrorDescription = "No se han podido calcular los totales por concepto para el resumen"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetSumary)
					If lErrorNumber = 0 Then
						sRowContents = "<BR />"
						sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">TOTAL " & asCompanies(2,lLastCompany-1) & "</FONT></B></TD></TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 01</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n01").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a01").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 04</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n04").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a04").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 05</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n05").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a05").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 06</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n06").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a06").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 07</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n07").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a07").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 08</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n08").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a08").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 11</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n11").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a11").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 44</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n44").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a44").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 7S</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n7s").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a7s").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto B2</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_nb2").Value,2,True,False,True)) & "</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_ab2").Value,2,True,False,True)) & "</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
						sRowContents = sRowContents & "</TABLE>"
						sRowContents = sRowContents & "<BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						Set oRecordsetSumary = Nothing
					End If
				End If

				If CLng(oRecordset.Fields("CompanyID").Value) <> CLng(lLastCompany) Then
					sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						sRowContents = sRowContents & "<TR>"
							sRowContents = sRowContents & "<TD WIDTH=""25%"">&nbsp;</TD>"
							sRowContents = sRowContents & "<TD WIDTH=""25%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>" & CStr(oRecordset.Fields("CompanyID").Value) & "</TD><TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>" & asCompanies(2,oRecordset.Fields("CompanyID").Value-1) & "</B></FONT></TD>"
						sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "</TABLE>"
					sRowContents = sRowContents & "<BR />"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lLastCompany = oRecordset.Fields("CompanyID").Value
				End If
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD WIDTH=""20%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>FECHA DE PAGO</FONT></B></TD>"
						sRowContents = sRowContents & "<TD WIDTH=""15%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("PaymentDate").Value), -1, -1, -1) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD WIDTH=""15%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>CONCEPTO</B></FONT></TD>"
						sRowContents = sRowContents & "<TD WIDTH=""15%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>CIFRAS RESUMEN</B></FONT></TD>"
						sRowContents = sRowContents & "<TD WIDTH=""15%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>CIFRAS SIAP</B></FONT></TD>"
						sRowContents = sRowContents & "<TD WIDTH=""20%"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>CIFRAS COSIF</B></FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 01</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n01").Value, 2, True, False, True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a01").Value, 2, True, False, True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Ingresos</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Income").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 04</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n04").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a04").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><INCOME /></FONT></TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 05</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n05").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a05").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Deducciones</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Deductions").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 06</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n06").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a06").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><DEDUCTIONS /></TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 07</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n07").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a07").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Líquido</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("NetIncome").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 08</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n08").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("cpt_a08").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NET_INCOME /></FONT></TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 11</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n11").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a11").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 44</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n44").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a44").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto B2</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_nB2").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_aB2").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 7S</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_n7s").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Cpt_a7s").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD><FONT FACE=""Arial"" SIZE=""2"">0</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				sRowContents = sRowContents & "<BR />"
				sRowContents = sRowContents & "<BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.MoveNext
			Loop
			sQuery = "Select Sum(Income) As Income, Sum(Deductions) As Deductions, Sum(NetIncome) As NetIncome, Sum(Cpt_n01) As Cpt_n01, Sum(Cpt_a01) As Cpt_a01, Sum(Cpt_n04) As Cpt_n04, Sum(Cpt_a04) As Cpt_a04, Sum(Cpt_n05) As Cpt_n05, Sum(Cpt_a05) As Cpt_a05, Sum(Cpt_n06) As Cpt_n06, Sum(Cpt_a06) As Cpt_a06, Sum(Cpt_n07) As Cpt_n07, Sum(Cpt_a07) As Cpt_a07, Sum(Cpt_n08) As Cpt_n08, Sum(Cpt_a08) As Cpt_a08, Sum(Cpt_n11) As Cpt_n11, Sum(Cpt_a11) As Cpt_a11, Sum(Cpt_n44) As Cpt_n44, Sum(Cpt_a44) As Cpt_a44, Sum(Cpt_n7s) As Cpt_n7s, Sum(Cpt_a7s) As Cpt_a7s, Sum(Cpt_nb2) As Cpt_nb2, Sum(Cpt_ab2) As Cpt_ab2 From DM_Hist_Binmar Where (CompanyID=" & lLastCompany & ")"
			sErrorDescription = "No se han podido calcular los totales por concepto para el resumen"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetSumary)
			If lErrorNumber = 0 Then
				sRowContents = "<BR />"
				sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				sRowContents = sRowContents & "<TR><TD COLSPAN=""5"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">TOTAL " & asCompanies(2,lLastCompany-1) & "</FONT></B></TD></TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 01</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n01").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a01").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 04</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n04").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a04").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 05</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n05").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a05").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 06</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n06").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a06").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 07</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n07").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a07").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 08</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n08").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a08").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 11</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n11").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a11").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 44</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n44").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a44").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto 7S</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_n7s").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_a7s").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto B2</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_nb2").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordsetSumary.Fields("Cpt_ab2").Value,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "</TR>"
			sRowContents = sRowContents & "</TABLE>"
			sRowContents = sRowContents & "<BR />"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			End If
			Set oRecordsetSumary = Nothing

		Call BuildReport1028_A(oRequest, oADODBConnection, sDocumentName, sFilePath, sFileName, sErrorDescription)

		End If
		If lErrorNumber = 0 Then
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		End If
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
		Else
			sErrorDescription = "No se ha podido generar el reporte de difras del bimestre"
		End If
	End If
	Set oRecordset = Nothing
	BuildReport1028 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1028_A(oRequest, oADODBConnection, sDocumentName, sFilePath, sFileName, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de cifras del bimestre)
'Inputs:  oRequest, oADODBConnection, sDocumentName, sFilePath, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1028_A"
	Dim sCondition
	Dim sQuery
	Dim asConcepts
	Dim afPayments1
	Dim afPayments2
	Dim fSubtotal1
	Dim fSubtotal2
	Dim iIndex
	Dim oRecordset
	Dim sRowContents
	
	asConcepts = Split("01,04,05,06,07,08,11,44,B2,7S",",")
	
	sQuery = "Select Sum(Cpt_n01) Cpt_n01, Sum(Cpt_n04) Cpt_n04, Sum(Cpt_n05) Cpt_n05, Sum(Cpt_n06) Cpt_n06, Sum(Cpt_n07) Cpt_n07, " & _
			"Sum(Cpt_n08) Cpt_n08, Sum(Cpt_n11) Cpt_n11, Sum(Cpt_n44) Cpt_n44, Sum(Cpt_nb2) Cpt_nb2, Sum(Cpt_n7s) Cpt_n7s " & _
			"From Dm_Hist_Binmar"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	afPayments1 = oRecordset.GetRows()

	sQuery = "Select Sum(Cpt_a01) Cpt_a01, Sum(Cpt_a04) Cpt_a04, Sum(Cpt_a05) Cpt_a05, Sum(Cpt_a06) Cpt_a06, Sum(Cpt_a07) Cpt_a07, " & _
			"Sum(Cpt_a08) Cpt_a08, Sum(Cpt_a11) Cpt_a11, Sum(Cpt_a44) Cpt_a44, Sum(Cpt_ab2) Cpt_ab2, Sum(Cpt_a7s) Cpt_a7s " & _
			"From Dm_Hist_Binmar"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	afPayments2 = oRecordset.GetRows()
	
	sRowContents = "<BR />"
		sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">TOTAL BIMESTRAL</FONT></B></TD></TR>"
			fSubtotal1 = 0
			fSubtotal2 = 0
	For iIndex = 0 To UBound(asConcepts)-1
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & asConcepts(iIndex) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(iIndex,0),2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments2(iIndex,0),2,True,False,True)) & "</FONT></TD>"
			sRowContents = sRowContents & "</TR>"
			fSubtotal1 = fSubtotal1 + afPayments1(iIndex,0)
			fSubtotal2 = fSubtotal2 + afPayments2(iIndex,0)
	Next
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD COLSPAN=""2"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD COLSPAN=""3""> <HR /></TD>"
			sRowContents = sRowContents & "</TR>"
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(fSubtotal1,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(fSubtotal2,2,True,False,True)) & "</FONT></TD>"
			sRowContents = sRowContents & "</TR>"
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & asConcepts(iIndex) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(iIndex,0),2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments2(iIndex,0),2,True,False,True)) & "</FONT></TD>"
			sRowContents = sRowContents & "</TR>"
			fSubtotal1 = fSubtotal1 + afPayments1(iIndex,0)
			fSubtotal2 = fSubtotal2 + afPayments2(iIndex,0)
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(fSubtotal1,2,True,False,True)) & "</FONT></TD>"
				sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(fSubtotal2,2,True,False,True)) & "</FONT></TD>"
			sRowContents = sRowContents & "</TR>"
		sRowContents = sRowContents & "</TABLE>"
		sRowContents = sRowContents & "<BR />"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

	
	sQuery = "Select ConceptShortName, Sum(ConceptAmount) Importe From Dm_Estr_Qna, Concepts Where (Dm_Estr_Qna.ConceptID = Concepts.ConceptID) And (Canceled = 0) Group By ConceptShortName Order By ConceptShortName"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	afPayments1=oRecordset.GetRows()
	sRowContents = "<BR />"
		sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">TOTAL BIMESTRAL</FONT></B></TD></TR>"
	For iIndex = 0 To UBound(afPayments1,2)
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & afPayments1(0,iIndex) & "</FONT></TD>"
				If StrComp(afPayments1(0,iIndex),"7S",vbBinaryCompare) <> 0 Then
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
				Else
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
				End If
			sRowContents = sRowContents & "</TR>"		
	Next
		sRowContents = sRowContents & "<TR><TD COLSPAN=""6""></TD></TR>"
		sRowContents = sRowContents & "</TABLE>"
		sRowContents = sRowContents & "<BR />"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

		sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">CIFRAS TOTALES DEL BIMESTRE</FONT></B></TD></TR>"
	For iIndex = 0 To UBound(afPayments1,2)
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & afPayments1(0,iIndex) & "</FONT></TD>"
				If StrComp(afPayments1(0,iIndex),"7S",vbBinaryCompare) <> 0 Then
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
				Else
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
				End If
			sRowContents = sRowContents & "</TR>"		
	Next
		sRowContents = sRowContents & "</TABLE>"
		sRowContents = sRowContents & "<BR />"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		
	sQuery = "Select ConceptShortName, Sum(ConceptAmount) Importe From Dm_Estr_Qna, Concepts Where (Dm_Estr_Qna.ConceptID = Concepts.ConceptID) And (Canceled = 1) Group By ConceptShortName Order By ConceptShortName"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
			afPayments2=oRecordset.GetRows()
			sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">REPORTE DE CIFRAS DE CANCELADOS DEL BIMESTRE</FONT></B></TD></TR>"
		For iIndex = 0 To UBound(afPayments2)
				sRowContents = sRowContents & "<TR>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
					sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & afPayments2(0,iIndex) & "</FONT></TD>"
					If StrComp(afPayments2(0,iIndex),"7S",vbBinaryCompare) <> 0 Then
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments2(1,iIndex),2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
					Else
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments2(1,iIndex),2,True,False,True)) & "</FONT></TD>"
					End If
				sRowContents = sRowContents & "</TR>"		
		Next
			sRowContents = sRowContents & "</TABLE>"
			sRowContents = sRowContents & "<BR />"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sQuery = "Select Concept.ConceptShortName, (Concept.Importe - Canceled.Importe) Amount From (Select ConceptShortName, Sum(ConceptAmount) Importe From Dm_Estr_Qna, Concepts Where (Dm_Estr_Qna.ConceptID = Concepts.ConceptID) And (Canceled = 0) Group By ConceptShortName) Concept, (Select ConceptShortName, Sum(ConceptAmount) Importe From Dm_Estr_Qna, Concepts Where (Dm_Estr_Qna.ConceptID = Concepts.ConceptID) And (Canceled = 1) Group By ConceptShortName) Canceled Where Concept.ConceptShortName = Canceled.ConceptShortName Order By Concept.ConceptShortName"
	Else
			sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">REPORTE DE CIFRAS DE CANCELADOS DEL BIMESTRE</FONT></B></TD></TR>"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No se encontraron registros cancelados</FONT></TD></TR>"
		sRowContents = sRowContents & "</TABLE>"
		sRowContents = sRowContents & "<BR />"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		sQuery = "Select ConceptShortName, Sum(ConceptAmount) Importe From Dm_Estr_Qna, Concepts Where (Dm_Estr_Qna.ConceptID = Concepts.ConceptID) And (Canceled = 0) Group By ConceptShortName Order By ConceptShortName"
	End If
	
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	afPayments1=oRecordset.GetRows()
	sRowContents = "<BR />"
		sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			sRowContents = sRowContents & "<TR>"
			sRowContents = sRowContents & "<TD COLSPAN=""2"">&nbsp;</TD>"
			sRowContents = sRowContents & "<TD COLSPAN=""4""><HR /></TD>"
			sRowContents = sRowContents & "</ TR>"
			sRowContents = sRowContents & "<TR><TD COLSPAN=""6"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">REPORTE DE CIFRAS TOTALES DEL BIMESTRE</FONT></B></TD></TR>"
	For iIndex = 0 To UBound(afPayments1,2)
			sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD WIDTH=""20%"">&nbsp;</TD>"
				sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto " & afPayments1(0,iIndex) & "</FONT></TD>"
				If StrComp(afPayments1(0,iIndex),"7S",vbBinaryCompare) <> 0 Then
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
				Else
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD>"
					sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(afPayments1(1,iIndex),2,True,False,True)) & "</FONT></TD>"
				End If
			sRowContents = sRowContents & "</TR>"		
	Next
			sRowContents = sRowContents & "<TR>"
			sRowContents = sRowContents & "<TD COLSPAN=""2"">&nbsp;</TD>"
			sRowContents = sRowContents & "<TD COLSPAN=""4""><HR /></TD>"
			sRowContents = sRowContents & "</ TR>"
		sRowContents = sRowContents & "</TABLE>"
		sRowContents = sRowContents & "<BR />"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)	
	
	Set oRecordset = Nothing
	BuildReport1028_A = lErrorNumber
	Err.Clear
End Function


Function BuildReport1029(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de total de incidencias registradas para el personal
'		  por cada centro de pago
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1029"
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
	Dim sGroupQuery

	Dim lCurrentPaymentCenterID
	Dim asStateNames
	Dim asAbsenceNames
	Dim asPath
	Dim iCount
	Dim aiAbscenceTotals
	Dim iIndex
	Dim sAbsenceShortName
	Dim bFirst
	Dim lTotal

	sQuery = "Select Employees.PaymentCenterID,EmployeesAbsencesLKP.AbsenceID,A.AbsenceShortName, A.AbsenceName,EmployeesAbsencesLKP.Active," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName,PaymentCenters.AreaName As PaymentCenterName,Zones.ZonePath," & _
			 " CompanyShortName,CompanyName,COUNT(*) As Total" & _
			 " From Employees, EmployeesAbsencesLKP, Absences As A, Justifications As J, Users, Areas, Areas As PaymentCenters, Jobs," & _
			 " Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.EmployeeID = EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.JustificationID = J.JustificationID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID = A.AbsenceID) And (EmployeesAbsencesLKP.AddUserID=Users.UserID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID < 100)"
			 
	sGroupQuery = " Group By Employees.PaymentCenterID,EmployeesAbsencesLKP.AbsenceID,A.AbsenceShortName, A.AbsenceName," & _
			 " EmployeesAbsencesLKP.Active,PaymentCenters.AreaCode,PaymentCenters.AreaName," & _
			 " Zones.ZonePath,CompanyShortName,CompanyName"

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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceShortName From Absences Where (AbsenceID>-1) And (AbsenceID<100) Order By AbsenceID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		asAbsenceNames = ""
		Do While Not oRecordset.EOF
			asAbsenceNames = asAbsenceNames & LIST_SEPARATOR & SizeText(CStr(CleanStringForHTML(oRecordset.Fields("AbsenceShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asAbsenceNames = Split(asAbsenceNames, LIST_SEPARATOR)
	End If
	aiAbscenceTotals = Split("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(aiAbscenceTotals)
		aiAbscenceTotals(iIndex) = 0
	Next

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "OcurredDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("OcurredDate", "EndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.Ocurred") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesAbsencesLKP.Ocurred", 1, 1, vbBinaryCompare) & "))"
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition & sCondition2 & sGroupQuery & " Order By PaymentCenters.AreaCode", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & sCondition & sCondition2 & sGroupQuery & " Order By PaymentCenters.AreaCode" & " -->" & vbNewLine
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
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1029.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					asConceptTitle = Split(aReportTitle(L_CONCEPT_ID_FLAGS), ";")
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
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
							If False Then
								sRowContents = "<BR /><B>TOTALES</B><BR />"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
								sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
								sRowContents = sRowContents & "<TD>TOTAL</TD>"
								sRowContents = sRowContents & "</FONT></TR>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								For iIndex = 0 To UBound(aiAbscenceTotals)
									lTotal = CInt(aiAbscenceTotals(iIndex))
									If lTotal > 0 Then
										sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
										sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
											sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
											sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
										sRowContents = sRowContents & "</FONT></TR>"
										lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									End If
								Next
								For iIndex = 0 To UBound(aiAbscenceTotals)
									aiAbscenceTotals(iIndex) = 0
								Next
								sRowContents = "</TABLE>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							End If
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
						sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
						sRowContents = sRowContents & "<TD>DESCRIPCION</TD>"
						sRowContents = sRowContents & "<TD>CLAVE CENTRO DE TRABAJO</TD>"
						sRowContents = sRowContents & "<TD>NOMBRE DEL CENTRO DE TRABAJO</TD>"
						sRowContents = sRowContents & "<TD>NUMERO DE REGISTROS</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("AbsenceShortName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("AbsenceName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PaymentCenterShortName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PaymentCenterName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("Total").Value) & "</TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + CInt(oRecordset.Fields("Total").Value)
					bFirst = True
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
					sRowContents = "</TABLE>"
					If False Then
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<BR /><B>TOTALES</B><BR />"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
						sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
						sRowContents = sRowContents & "<TD>TOTAL</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						For iIndex = 0 To UBound(aiAbscenceTotals)
							lTotal = CInt(aiAbscenceTotals(iIndex))
							If lTotal > 0 Then
								sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
								sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
									sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
									sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
								sRowContents = sRowContents & "</FONT></TR>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							End If
						Next
						For iIndex = 0 To UBound(aiAbscenceTotals)
							aiAbscenceTotals(iIndex) = 0
						Next
						sRowContents = "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				End If
				oRecordset.Close
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1029 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1030(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de incidencias registradas para el personal filtrado por
'         número de empleado, áreas y período específico
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1030"
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

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asAbsenceNames
	Dim asPath
	Dim iCount
	Dim aiAbscenceTotals
	Dim aiAbscenceGrandTotals
	Dim iIndex
	Dim sAbsenceShortName
	Dim bFirst
	Dim lTotal
	Dim lTotalForArea
	Dim lTotalForReport
	Dim iMin
	Dim iMax

	sQuery = "Select EmployeesAbsencesLKP.EmployeeID, Employees.EmployeeNumber, Employees.PaymentCenterID, EmployeesAbsencesLKP.AbsenceID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," & _
			 " EmployeesAbsencesLKP.AppliedDate, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, EmployeesAbsencesLKP.RegistrationDate, EmployeesAbsencesLKP.VacationPeriod, EmployeesAbsencesLKP.DocumentNumber, EmployeesAbsencesLKP.AbsenceHours," & _
			 " EmployeesAbsencesLKP.JustificationID, J.JustificationShortName, EmployeesAbsencesLKP.Reasons, EmployeesAbsencesLKP.Removed, EmployeesAbsencesLKP.JustificationID As AbsenceJustified, EmployeesAbsencesLKP.Active, Employees.JourneyID," & _
			 " EmployeesAbsencesLKP.AppliedRemoveDate, A.AbsenceShortName, A.AbsenceName, A.IsJustified, A.JustificationID As WithJustification, Users.UserLastName + ' ' + Users.UserName As UserFullName," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName," & _
			 " AnotherEmployeesAbsencesLKP.AbsenceID As AnotherAbsenceID, AnotherEmployeesAbsencesLKP.AbsenceHours As AnotherAbsenceHours" & _
			 " From Employees, EmployeesAbsencesLKP, EmployeesAbsencesLKP As AnotherEmployeesAbsencesLKP, Absences As A, Justifications As J, Users, Areas, Areas As PaymentCenters, Jobs," & _
			 " Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.JustificationID=J.JustificationID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID=A.AbsenceID) And (EmployeesAbsencesLKP.AddUserID=Users.UserID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (EmployeesAbsencesLKP.AbsenceID In (10,11,12,13,14,16,17,82,83,84,85,86,87,29,30,31,32,33,34,35,37,38))" & _
			 " And (EmployeesAbsencesLKP.EmployeeID=AnotherEmployeesAbsencesLKP.EmployeeID)" & _
			 " And (AnotherEmployeesAbsencesLKP.AbsenceID IN (201,202))" & _
			 " And (((EmployeesAbsencesLKP.OcurredDate>=AnotherEmployeesAbsencesLKP.OcurredDate) And (EmployeesAbsencesLKP.OcurredDate<=AnotherEmployeesAbsencesLKP.EndDate))" & _
			 " Or ((EmployeesAbsencesLKP.EndDate>=AnotherEmployeesAbsencesLKP.OcurredDate) And (EmployeesAbsencesLKP.EndDate<=AnotherEmployeesAbsencesLKP.EndDate))" & _
			 " Or ((EmployeesAbsencesLKP.EndDate>=AnotherEmployeesAbsencesLKP.OcurredDate) And (EmployeesAbsencesLKP.OcurredDate<=AnotherEmployeesAbsencesLKP.EndDate)))"

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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(AbsenceID) As Max From Absences", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value)
		End If
	End If
	For iMin = 0 To iMax
		asAbsenceNames = asAbsenceNames & LIST_SEPARATOR & ""
		aiAbscenceTotals = aiAbscenceTotals & LIST_SEPARATOR & "0"
		aiAbscenceGrandTotals = aiAbscenceGrandTotals & LIST_SEPARATOR & "0"
	Next
	asAbsenceNames = Split(asAbsenceNames, LIST_SEPARATOR)
	aiAbscenceTotals = Split(aiAbscenceTotals, LIST_SEPARATOR)
	aiAbscenceGrandTotals = Split(aiAbscenceGrandTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiAbscenceTotals(iIndex) = 0
		aiAbscenceGrandTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, AbsenceShortName From Absences Where (AbsenceID>-1) And (AbsenceID<100) Order By AbsenceID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asAbsenceNames(CInt(oRecordset.Fields("AbsenceID").Value)) = SizeText(CStr(CleanStringForHTML(oRecordset.Fields("AbsenceShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "OcurredDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("OcurredDate", "EndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.Ocurred") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesAbsencesLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesAbsencesLKP.Ocurred", 1, 1, vbBinaryCompare) & "))"
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
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
					sHeaderContents = Replace(sHeaderContents, "<TITLE />", CleanStringForHTML("Reporte de Incidencias con horas extras y primas dominicales el mismo día."))
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
					lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				End If
				iCount = 0
				lCurrentPaymentCenterID = -2
				lTotalForReport = 0
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
							sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
							sRowContents = sRowContents & "<TD>TOTAL</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							lTotalForArea = 0
							For iIndex = 0 To UBound(aiAbscenceTotals)
								lTotal = CInt(aiAbscenceTotals(iIndex))
								If lTotal > 0 Then
									lTotalForArea = lTotalForArea + lTotal
									lTotalForReport = lTotalForReport + lTotal
									sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
										sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
									sRowContents = sRowContents & "</FONT></TR>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							Next
							For iIndex = 0 To UBound(aiAbscenceTotals)
								aiAbscenceTotals(iIndex) = 0
							Next
							sRowContents = "</TABLE>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							sRowContents = "<BR /><B>REGISTROS TOTALES POR CENTRO DE TRABAJO: " & lTotalForArea & "</B><BR />"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
							sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
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
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de ocurrencia</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de término</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Periodo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave del centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Cantidad</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Estatus</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Justificación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Q. de aplicación de la justificación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No. de documento</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del usuario</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">H.Extra/P.Dominical</FONT></TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CInt(oRecordset.Fields("AbsenceID").Value)
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeFullName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("=T(""" & CStr(oRecordset.Fields("AbsenceShortName").Value)) & """)</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "</FONT></TD>"
						If (Not VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) Or ((CInt(oRecordset.Fields("JourneyID").Value)=21) Or (CInt(oRecordset.Fields("JourneyID").Value)=22) Or (CInt(oRecordset.Fields("JourneyID").Value)=23))) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>"
						Else
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("A la fecha") & "</FONT></TD>"
							Else
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
							End If
						End If
						If CInt(oRecordset.Fields("VacationPeriod").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("NA") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Left(CStr(oRecordset.Fields("VacationPeriod").Value), Len("0000")) & "-" & Right(CStr(oRecordset.Fields("VacationPeriod").Value), Len("0"))) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("Aplicada") & "</FONT></TD>"
						ElseIf CInt(oRecordset.Fields("Active").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("En proceso") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("Cancelada") & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("JustificationShortName").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("JustificationID").Value) = -1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("NA") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedRemoveDate").Value)) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("AnotherAbsenceID").Value) = 201 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AnotherAbsenceHours").Value) & " H. Extras") & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("AnotherAbsenceHours").Value) & " P. Dominical") & "</FONT></TD>"
						End If
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					If CInt(oRecordset.Fields("JustificationID").Value) <> -1 Then
						aiAbscenceTotals(CInt(oRecordset.Fields("JustificationID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("JustificationID").Value)) + 1
						aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) + 1
					Else
						aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
						aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
					End If
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
					sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
					sRowContents = sRowContents & "<TD>TOTAL</TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lTotalForArea = 0
					For iIndex = 0 To UBound(aiAbscenceTotals)
						lTotal = CInt(aiAbscenceTotals(iIndex))
						If lTotal > 0 Then
							lTotalForArea = lTotalForArea + lTotal
							lTotalForReport = lTotalForReport + lTotal
							sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
					For iIndex = 0 To UBound(aiAbscenceTotals)
						aiAbscenceTotals(iIndex) = 0
					Next
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					sRowContents = "<BR /><B>REGISTROS TOTALES POR CENTRO DE TRABAJO: " & lTotalForArea & "</B><BR />"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				End If
				sRowContents = "<BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>TOTALES DEL REPORTE</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
				sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
				sRowContents = sRowContents & "<TD>TOTAL</TD>"
				sRowContents = sRowContents & "</FONT></TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				For iIndex = 0 To UBound(aiAbscenceGrandTotals)
					lTotal = CInt(aiAbscenceGrandTotals(iIndex))
					If lTotal > 0 Then
						sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
							sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				Next
				For iIndex = 0 To UBound(aiAbscenceGrandTotals)
					aiAbscenceGrandTotals(iIndex) = 0
				Next
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>REGISTROS TOTALES DEL REPORTE: " & lTotalForReport & "</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.Close
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
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1030 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1031(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de altas, bajas y cambios del bimestre)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1031"
	Dim asReports
	Dim asTitles
	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sQuery
	Dim sReports
	Dim sReportType
	Dim sRowContents
	Dim sTitles
	Dim lErrorNumber
	Dim oRecordset
	Dim iIndex
	Dim iReportType

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()
	If (oRequest("ReportType").Item = 0) Then
		lErrorNumber = BuildReport1031_A(oRequest, oADODBConnection, sFilePath, 2, sDate, sErrorDescription)
		lErrorNumber = BuildReport1031_B(oRequest, oADODBConnection, sFilePath, 3, sDate, sErrorDescription)
		lErrorNumber = BuildReport1031_C(oRequest, oADODBConnection, sFilePath, 4, sDate, sErrorDescription)
		lErrorNumber = BuildReport1031_C(oRequest, oADODBConnection, sFilePath, 5, sDate, sErrorDescription)
		lErrorNumber = BuildReport1031_C(oRequest, oADODBConnection, sFilePath, 6, sDate, sErrorDescription)
		lErrorNumber = BuildReport1031_C(oRequest, oADODBConnection, sFilePath, 7, sDate, sErrorDescription)
	Else
		iReportType = oRequest("ReportType").Item
		Select case iReportType
			Case 2
				lErrorNumber = BuildReport1031_A(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
			Case 3
				lErrorNumber = BuildReport1031_B(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
			Case 4,5,6,7
				lErrorNumber = BuildReport1031_C(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
		End Select
	End If
	If lErrorNumber = 0 Then
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
	End If
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
	Set oRecordset = Nothing
	BuildReport1031 = lErrorNumber
	Err.Clear
End Function


Function BuildReport1031_A(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de altas del bimestre)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1031_A"
	Dim sCondition
	Dim sDocumentName
	Dim sHeaderContents
	Dim sQuery
	Dim sRowContents
	Dim lErrorNumber
	Dim oRecordset

	sQuery = "Select * From dm_padron_banamex Where "
	sCondition = "(Status = " & iReportType & ")"
	If Len(oRequest("EmployeeNumbers").Item) > 0 Then sCondition = sCondition & " And (EmployeeID In (" & oRequest("EmployeeNumbers").Item & ")) "
	If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " And (u_version = " & oRequest("CompanyID").Item & ") "
	If Len(oRequest("ZoneID").Item) > 0 Then sCondition = sCondition & " And (State = " & oRequest("ZoneID").Item & ") "
	sQuery = sQuery & sCondition & " Order By EmployeeID Asc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
		sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_Altas.txt"
		sHeaderContents = GetFileContents(Server.MapPath("Templates\altas_bim.txt"), sErrorDescription)
		lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		Do While Not oRecordset.EOF
			sRowContents = "A"
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("RFC").Value, "", " ")) & String(13," "),13)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("CURP").Value, "", " ")) & String(18," "),18)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("SocialSecurityNumber").Value, "", " ")) & String(10," "),10)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeLastName").Value) & String(40," "),40)
			If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
				sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("EmployeeLastName2").Value, "", " ")) & String(40," "),40)
			Else
				sRowContents = sRowContents & Left(String(40," "),40)
			End If
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeName").Value) & String(40," "),40)
			sRowContents = sRowContents & CStr(Right("00000" & CStr(u_version),5))
			sRowContents = sRowContents & CStr(Left(oRecordset.Fields("CT").Value & "0000000000",20))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("BirthDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00" & CStr(BirhtState),2))
			sRowContents = sRowContents & CStr(oRecordset.Fields("GenderShortName").Value)
			sRowContents = sRowContents & CStr(oRecordset.Fields("MaritalStatusID").Value)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("Address").Value) & String(60," "),60)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("Colony").Value) & String(30," "),30)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("City").Value) & String(30," "),30)
			sRowContents = sRowContents & CStr(Right("00000" & CStr(oRecordset.Fields("ZipZone").Value),5))
			sRowContents = sRowContents & CStr(Right("00" & CStr(oRecordset.Fields("State").Value),2))
			sRowContents = sRowContents & CStr(oRecordset.Fields("Nombram").Value)
			sRowContents = sRowContents & CStr(Right("0000000000" & CStr(oRecordset.Fields("EmployeeID").Value),10))
			sRowContents = sRowContents & CStr(Right("000" & CStr(oRecordset.Fields("ICEFA").Value),3))
			sRowContents = sRowContents & CStr(oRecordset.Fields("Afore").Value)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("JoinDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("CotDate").Value),8))
			sRowContents = sRowContents & CStr(oRecordset.Fields("Fovi").Value)
			sRowContents = sRowContents & CStr(Right("000" & CStr(oRecordset.Fields("WorkingDays").Value),3))
			sRowContents = sRowContents & CStr(Right("000" & CStr(oRecordset.Fields("InabilityDays").Value),3))
			sRowContents = sRowContents & CStr(Right("000" & CStr(oRecordset.Fields("AbsenceDays").Value),3))
			sRowContents = sRowContents & CStr(Right("0000000" & CStr(oRecordset.Fields("Salary").Value),7))
			sRowContents = sRowContents & CStr(Right("0000000" & CStr(oRecordset.Fields("FullPay").Value),7))
			sRowContents = sRowContents & CStr(Right("0000000" & CStr(oRecordset.Fields("Salary_v").Value),7))
			sRowContents = sRowContents & CStr(oRecordset.Fields("EmployeeContributions").Value)
			sRowContents = sRowContents & CStr(Right("0000000" & CStr(oRecordset.Fields("EmployeeContributionsAmount").Value),7))
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			oRecordset.MoveNext
			If lErrorNumber <> 0 Then Exit Do
		Loop
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1031_A = lErrorNumber
	Err.Clear
End Function

Function BuildReport1031_B(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de bajas del bimestre)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1031_B"
	Dim sCondition
	Dim sDocumentName
	Dim sHeaderContents
	Dim sQuery
	Dim sRowContents
	Dim lErrorNumber
	Dim oRecordset

	sQuery = "Select * From dm_padron_banamex Where "
	sCondition = "(Status = " & iReportType & ")"
	If Len(oRequest("EmployeeNumbers").Item) > 0 Then sCondition = sCondition & " And (EmployeeID In (" & oRequest("EmployeeNumbers").Item & ")) "
	If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " And (u_version = " & oRequest("CompanyID").Item & ") "
	If Len(oRequest("ZoneID").Item) > 0 Then sCondition = sCondition & " And (State = " & oRequest("ZoneID").Item & ") "
	sQuery = sQuery & sCondition & " Order By EmployeeID Asc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
		sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_Bajas.txt"
		sHeaderContents = GetFileContents(Server.MapPath("Templates\altas_bim.txt"), sErrorDescription)
		lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		Do While Not oRecordset.EOF
			sRowContents = "B"
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("RFC").Value, "", " ")) & String(13," "),13)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("CURP").Value, "", " ")) & String(18," "),18)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("SocialSecurityNumber").Value, "", " ")) & String(10," "),10)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeLastName").Value) & String(40," "),40)
			If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
				sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("EmployeeLastName2").Value, "", " ")) & String(40," "),40)
			Else
				sRowContents = sRowContents & Left(String(40," "),40)
			End If
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeName").Value) & String(40," "),40)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("BirthDate").Value),8))
			sRowContents = sRowContents & Left(oRecordset.Fields("GenderShortName").Value & String(2," "),2)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("JoinDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("CotDate").Value),8))
			sRowContents = sRowContents & CStr(oRecordset.Fields("mot_baja").Value)
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			oRecordset.MoveNext
			If lErrorNumber <> 0 Then Exit Do
		Loop
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1031_B = lErrorNumber
	Err.Clear
End Function

Function BuildReport1031_C(oRequest, oADODBConnection, sFilePath, iReportType, sDate, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de cambios del bimestre)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1031_C"
	Dim sCondition
	Dim sDocumentName
	Dim sHeaderContents
	Dim sQuery
	Dim sRowContents
	Dim lErrorNumber
	Dim oRecordset

	sQuery = "Select Distinct ant.rfc rfcAnt,ant.curp curpAnt,ant.SocialSecurityNumber ssnAnt,ant.EmployeeLastName lastNameAnt,ant.EmployeeLastName2 lastName2Ant,ant.EmployeeName nameAnt,act.* From dm_padron_banamex act, dm_update_padron_banamex ant Where (act.EmployeeID = ant.EmployeeID)"
	sCondition = " And (act.Status = " & iReportType & ")"
	If Len(oRequest("EmployeeNumbers").Item) > 0 Then sCondition = sCondition & " And (act.EmployeeID In (" & oRequest("EmployeeNumbers").Item & ")) "
	If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " And (act.u_version = " & oRequest("CompanyID").Item & ") "
	If Len(oRequest("ZoneID").Item) > 0 Then sCondition = sCondition & " And (act.State = " & oRequest("ZoneID").Item & ") "
	sQuery = sQuery & sCondition & " Order By act.EmployeeID Asc"
	
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
		'cambios_sar_abrecierre
		'cambios_sar_dias
		'cambios_sar_sdos
		'cambios_sar_voluntaria
		Select Case iReportType
			Case 4
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & sDate & "_abrecierre.txt"
			Case 5
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & sDate & "_sdo.txt"
			Case 6
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & sDate & "_dias.txt"
			Case 7
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & sDate & "_Voluntaria.txt"
		End Select
		sHeaderContents = GetFileContents(Server.MapPath("Templates\altas_bim.txt"), sErrorDescription)
		lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		Do While Not oRecordset.EOF
			sRowContents = "M"
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("rfcAnt").Value, "", " ")) & String(13," "),13)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("curpAnt").Value, "", " ")) & String(18," "),18)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("nssAnt").Value, "", " ")) & String(10," "),10)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("LastNameAnt").Value) & String(40," "),40)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("LastName2Ant").Value, "", " ")) & String(40," "),40)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("NameAnt").Value) & String(40," "),40)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("rfc").Value, "", " ")) & String(13," "),13)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("curp").Value, "", " ")) & String(18," "),18)
			sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("SocialSecurityNumber").Value, "", " ")) & String(10," "),10)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("ct").Value) & String(20," "),20)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("u_version").Value),8))
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeLastName").Value) & String(40," "),40)
			If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
				sRowContents = sRowContents & Left(CStr(Replace(oRecordset.Fields("EmployeeLastName2").Value, "", " ")) & String(40," "),40)
			Else
				sRowContents = sRowContents & Left(String(40," "),40)
			End If
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("EmployeeName").Value) & String(40," "),40)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("BirthDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("BirthState").Value),2))
			sRowContents = sRowContents & CStr(Right("00" & oRecordset.Fields("GenderShortName").Value,2))
			sRowContents = sRowContents & CStr(Right("00" & CStr(oRecordset.Fields("MaritalStatusID").Value),2))
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("Address").Value) & String(60," "),60)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("Colony").Value) & String(30," "),30)
			sRowContents = sRowContents & Left(CStr(oRecordset.Fields("City").Value) & String(30," "),30)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("ZipZone").Value),5))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("State").Value),2))
			sRowContents = sRowContents & CStr(oRecordset.Fields("nombram").Value)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("EmployeeID").Value),10))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("ICEFA").Value),3))
			sRowContents = sRowContents & CStr(oRecordset.Fields("afore").Value)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("JoinDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("CotDate").Value),8))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("LastUpdateDate").Value),8))
			sRowContents = sRowContents & CStr(oRecordset.Fields("Fovi").Value)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("WorkingDays").Value),3))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("InabilityDays").Value),3))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("AbsenceDays").Value),3))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("Salary").Value),5))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("salary_v").Value),5))
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("FullPay").Value),5))
			sRowContents = sRowContents & CStr(oRecordset.Fields("EmployeeContributions").Value)
			sRowContents = sRowContents & CStr(Right("00000000" & CStr(oRecordset.Fields("EmployeeContributionsAmount").Value),5))
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			oRecordset.MoveNext
			If lErrorNumber <> 0 Then Exit Do
		Loop
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1031_C = lErrorNumber
	Err.Clear
End Function

Function BuildReport1032(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reportes de dispersión por UA y Empresa
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1032"
	Dim asCompanies
	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sQuery
	Dim sRowContents
	Dim lErrorNumber
	Dim oRecordset
	Dim iIndex
	Dim iCompany
	Dim fSar
	Dim fEntityCv
	Dim fEmployeeCv
	Dim fFoviAmount
	Dim fEmployeeSaving
	Dim fEntitySaving

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	sQuery = "Select CompanyID, CompanyName From Companies Where CompanyID > 0"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asCompanies = oRecordset.GetRows

	If lErrorNumber = 0 Then
		sDate = GetSerialNumberForDate("")
		sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
		lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
		sFilePath = sFilePath & "\"
		sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		For iCompany = 0 To UBound(asCompanies,2)
			fSar = 0
			fEntityCv = 0
			fEmployeeCv = 0
			fFoviAmount = 0
			fEmployeeSaving = 0
			fEntitySaving = 0
			sQuery = "Select SUBSTRING (DM_PADRON_BANAMEX.CT,1,2) Zone, SUM(SAR) Sar, SUM(entityCV) entityCV, SUM(EmployeeCV) EmployeeCV, SUM(foviAmount) FoviAmount, SUM(employeeSaving) employeeSaving, SUM(EntitySaving) entitySaving From DM_APORT_SAR, DM_PADRON_BANAMEX Where (CompanyID = " & asCompanies(0,iCompany) & ") And (DM_APORT_SAR.CURP = DM_PADRON_BANAMEX.CURP) Group by SUBSTRING (DM_PADRON_BANAMEX.CT,1,2) Order by 1 Asc"			
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If Not oRecordset.EOF Then
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_0" & iCompany & ".htm"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1028_2.htm"), sErrorDescription)
				sHeaderContents = Replace(sHeaderContents, "<COMPANY_NAME />", asCompanies(1,iCompany))
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = sRowContents & "<BR />"
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						If Len(oRecordset.Fields("Zone").Value) < 2 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""ARIAL"" SIZE=""2"">0" & oRecordset.Fields("Zone").Value & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""ARIAL"" SIZE=""2"">" & oRecordset.Fields("Zone").Value & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Sar").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("entityCV").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("EmployeeCV").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("FoviAmount").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("EmployeeSaving").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("EntitySaving").Value,2,True,False,True)) & "</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					fSar = fSar + oRecordset.Fields("Sar").Value
					fEntityCv = fEntityCv + oRecordset.Fields("entityCV").Value
					fEmployeeCv = fEmployeeCv + oRecordset.Fields("EmployeeCV").Value
					fFoviAmount = fFoviAmount + oRecordset.Fields("FoviAmount").Value
					fEmployeeSaving = fEmployeeSaving + oRecordset.Fields("EmployeeSaving").Value
					fEntitySaving = fEntitySaving + oRecordset.Fields("EntitySaving").Value
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If lErrorNumber <> 0 Then Exit Do
				Loop
				sRowContents = sRowContents & "</TABLE>"
				sRowContents = sRowContents & "<HR />"
				sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fSar,2,True,False,True)) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fEntityCv,2,True,False,True)) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fEmployeeCv,2,True,False,True)) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fFoviAmount,2,True,False,True)) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fEmployeeSaving,2,True,False,True)) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""ARIAL"" SIZE=""2""><B>" & CStr(FormatNumber(fEntitySaving,2,True,False,True)) & "</B></FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			End If
		Next
		If lErrorNumber = 0 Then
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		End If
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
	BuildReport1032 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1033(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Reporte de aportaciones)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1033"
	Dim asFields
	Dim asTitles
	Dim asDocuments
	Dim asCompanies
	Dim sCondition
	Dim sCompanies
	Dim sDate
	Dim sDocuments
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sFields
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sQuery
	Dim sRowContents
	Dim sTitles
	Dim lErrorNumber
	Dim lPeriodID
	Dim lTotalAmount
	Dim lTotalCount
	Dim oRecordset
	Dim iIndex
	Dim jIndex
	Dim kIndex

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	sQuery = "Select CompanyID, CompanyName From Companies Where CompanyID > 0"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asCompanies = oRecordset.GetRows

	sTitles = "voluntarias patronales,voluntarias del trabajador,cesantía y vejez del trabajador,cesantía y vejez patronales,al FOVISSSTE,al SAR"
	asTitles = Split(sTitles, ",")
	
	sFields = "entitySaving,employeeSaving,employeeCV,entityCV,foviAmount,sar"
	asFields = Split(sFields, ",")

	sQuery = "Select Distinct PeriodID From DM_Hist_Nomsar"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lPeriodID = oRecordset.Fields("PeriodID").Value
	
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()
	For jIndex = 0 To UBound(asCompanies,2)
		For iIndex = 0 To UBound(asTitles)
			lTotalAmount = 0
			lTotalCount = 0
			sQuery = "Select CT, COUNT(" & asFields(iIndex) & ") Casos, SUM(" & asFields(iIndex) & ") Monto From DM_APORT_SAR Where CompanyID = " & asCompanies(0,jIndex) & " Group By CT Having SUM(" & asFields(iIndex) & ") > 0 Order By CT Asc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_R" & iIndex & "_" & asCompanies(0,jIndex) & ".xls"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1028_3C.htm"), sErrorDescription)
				sHeaderContents = Replace(sHeaderContents, "<COMPANY_NAME />", asCompanies(1,jIndex))
				sHeaderContents = Replace(sHeaderContents, "<PERIOD_ID />", lPeriodID)
				sHeaderContents = Replace(sHeaderContents, "<REPORT_NAME />", asTitles(iIndex))
				sHeaderContents = Replace(sHeaderContents, "<PRINTING_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<PRINTING_HOUR />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<BR />"
				sRowContents = sRowContents & "<TABLE WIDTH=""50%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
						sRowContents = "<TR>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""ARIAL"" SIZE=""2"">" & CStr(oRecordset.Fields("CT").Value) & "</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("Casos").Value) & "</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(FormatNumber(oRecordset.Fields("Monto").Value,2,True,False,True)) & "</FONT></TD>"
						sRowContents = sRowContents & "</ TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lTotalCount = lTotalCount + oRecordset.Fields("Casos").Value
					lTotalAmount = lTotalAmount + oRecordset.Fields("Monto").Value
					oRecordset.MoveNext
					If lErrorNumber <> 0 Then Exit Do
				Loop
				sRowContents = "</TABLE>"
				sRowContents = sRowContents & "<LINE />"
				sRowContents = sRowContents & "<TABLE WIDTH=""50%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
				sRowContents = sRowContents & "</HR>"
				sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""ARIAL"" SIZE=""2""><B>TOTALES</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><B>" & CStr(lTotalCount) & "</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><B>" & CStr(FormatNumber(lTotalAmount,2,True,False,True)) & "</B></FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"	
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				End If
			End If
		Next
	Next

	For jIndex = 0 To UBound(asCompanies,2)
		For iIndex = 0 To UBound(asTitles)
			lTotalAmount = 0
			lTotalCount = 0
			sQuery = "Select CT, COUNT(" & asFields(iIndex) & ") Casos, SUM(" & asFields(iIndex) & ") Monto From DM_APORT_SAR Where CompanyID = " & asCompanies(0,jIndex) & " Group By CT Having SUM(" & asFields(iIndex) & ") > 0 Order By CT Asc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_R" & iIndex & "_" & asCompanies(0,jIndex) & ".txt"
				'sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1028_3C.htm"), sErrorDescription)
				'lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				'lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
							sRowContents = Right(String(10,"0") & CStr(oRecordset.Fields("CT").Value),10)
							sRowContents = sRowContents & Right(String(8,"0") & CStr(oRecordset.Fields("Casos").Value),8)
							sRowContents = sRowContents & Right(String(16,"0") & CStr(oRecordset.Fields("Monto").Value*100),16)
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If lErrorNumber <> 0 Then Exit Do
				Loop
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				End If
			End If
		Next
	Next

	If lErrorNumber = 0 Then
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
	End If
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

	Set oRecordset = Nothing
	BuildReport1033 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1034(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  (Control y distribución de comprobantes de abono
'		  en cuenta de trabajadores)
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1034"
	Dim asQuery
	Dim asStates
	Dim asTitles
	Dim sCity
	Dim sCondition
	Dim sCt
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sQuery
	Dim sRowContents
	Dim sState
	Dim lErrorNumber
	Dim oRecordset
	Dim iIndex
	Dim lPeriodID
	Dim lTotalAmount
	Dim lTotalCount

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCity = "."
	sCt = "."

	sQuery = "Select ZoneID, ZoneName From Zones Where (parentid = -1) And (ZoneID <> -1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asStates = oRecordset.GetRows()

	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()

	For iIndex = 0 To UBound(asStates,2)
        sQuery = "Select CT, EmployeeID, Z.ZoneName, '" & asStates(1,iIndex) & "', EmployeeLastName, EmployeeLastName2, EmployeeName From Dm_Aport_Sar Das, (Select ZoneID, ZoneCode, ZoneName From Zones Where ZonePath Like ',-1," & asStates(0,iIndex) & ",%') Z Where (Das.zonecode = Z.ZoneID) and (das.zonecode IS NOT NULL or Das.zonecode<> '') Order by Z.ZoneName, CT Asc,  EmployeeLastName Asc, EmployeeLastName2 Asc, EmployeeName Asc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sDocumentName = sFilePath & "Rep_" & asStates(1,iIndex) & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".htm"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1028_3A.htm"), sErrorDescription)
				sHeaderContents = Replace(sHeaderContents, "<PRINTING_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
				sHeaderContents = Replace(sHeaderContents, "<PRINTING_HOUR />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				Do While Not oRecordset.EOF
                    If  strComp(sCt,oRecordset.Fields("CT").Value,vbBinaryCompare) <> 0 or strComp(sCity,oRecordset.Fields("ZoneName").Value,vbBinaryCompare) <> 0 Then
						If lErrorNumber = 0 Then
							sRowContents = "</TABLE> <HR />"
							sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
								sRowContents = sRowContents & "<TR>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>ADSCRIPCIÓN</B></FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & oRecordset.Fields("CT").Value & "</FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>POBLACIÓN</B></FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & oRecordset.Fields("ZoneName").Value & "</FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>ESTADO</B></FONT></TD>"
									If Not IsNull(oRecordset.Fields("Estado").Value) Then
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & asStates(1,iIndex) & "</FONT></TD>"
									Else
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & oRecordset.Fields("ZoneName").Value & "</FONT></TD>"
									End If
							sRowContents = sRowContents & "</TABLE>"
							sRowContents = sRowContents & "<HR />"
							sRowContents = sRowContents & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING="""" CELLSPACING=""0"">"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							
						End If
					End If
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""LEFT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("RFC").Value) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD ALIGN=""LEFT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""LEFT""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">|__________________________________|</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">       /    /    </FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">|__________________________________|</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

                    sCt = oRecordset.Fields("CT").Value
                    sCity = oRecordset.Fields("ZoneName").Value

					oRecordset.MoveNext
					If lErrorNumber <> 0 Then Exit Do
				Loop
				sRowContents = sRowContents & "</TABLE>"
				If lErrorNumber = 0 Then
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No existen registros que cumplan con los criterios de la búsqueda."
			End If
		End If
	Next
	
	If lErrorNumber = 0 Then
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
	End If
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

	Set oRecordset = Nothing
	BuildReport1034 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1035(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reportes para el ejercicio bimestral del SAR.
'		  Relación de empleados dados de alta por comparación
'		  de nóminas y generación de archivo con referencias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1035"
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sQuery
	Dim sRowContents
	Dim sState
	Dim lErrorNumber
	Dim oRecordset
	Dim iIndex

	sQuery = "Select * From Dm_Padron_Banamex Where Status=2 Order By EmployeeID Asc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			sDocumentName = sFilePath & "Rep_Altas" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"

			sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "<TR>"
				sRowContents = sRowContents & "<TD>Empresa</TD>"
				sRowContents = sRowContents & "<TD>No. Empleado</TD>"
				sRowContents = sRowContents & "<TD>RFC</TD>"
				sRowContents = sRowContents & "<TD>CURP</TD>"
				sRowContents = sRowContents & "<TD>Núm. Seguro Social</TD>"
				sRowContents = sRowContents & "<TD>Apellido paterno</TD>"
				sRowContents = sRowContents & "<TD>apellido materno</TD>"
				sRowContents = sRowContents & "<TD>Nombre</TD>"
				sRowContents = sRowContents & "<TD>CT</TD>"
				sRowContents = sRowContents & "<TD>Fecha de nacimiento</TD>"
				sRowContents = sRowContents & "<TD>Estado de nacimiento</TD>"
				sRowContents = sRowContents & "<TD>Género</TD>"
				sRowContents = sRowContents & "<TD>Fecha de contratación</TD>"
				sRowContents = sRowContents & "<TD>Fecha de cotización</TD>"
				sRowContents = sRowContents & "<TD>Periodo</TD>"
				sRowContents = sRowContents & "<TD>Estado civil</TD>"
				sRowContents = sRowContents & "<TD>Dirección</TD>"
				sRowContents = sRowContents & "<TD>Colonia</TD>"
				sRowContents = sRowContents & "<TD>Ciudad</TD>"
				sRowContents = sRowContents & "<TD>Código postal</TD>"
				sRowContents = sRowContents & "<TD>Estado</TD>"
				sRowContents = sRowContents & "<TD>Nombramiento</TD>"
				sRowContents = sRowContents & "<TD>Afore</TD>"
				sRowContents = sRowContents & "<TD>ICEFA</TD>"
				sRowContents = sRowContents & "<TD>Núm. Control Interno</TD>"
				sRowContents = sRowContents & "<TD>Salario</TD>"
				sRowContents = sRowContents & "<TD>Salario_v</TD>"
				sRowContents = sRowContents & "<TD>Salario integrado</TD>"
				sRowContents = sRowContents & "<TD>Días laborados</TD>"
				sRowContents = sRowContents & "<TD>Contribución del empleado</TD>"
				sRowContents = sRowContents & "<TD>Monto</TD>"
				sRowContents = sRowContents & "<TD>Fecha de inicio</TD>"
				sRowContents = sRowContents & "<TD>Fecha fin</TD>"
			sRowContents = sRowContents & "</TR>"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

			Do While Not oRecordset.EOF
				sRowContents = "<TR>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("u_version").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("rfc").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("curp").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</TD>"
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</TD>"
					Else
						sRowContents = sRowContents & "<TD>&nbsp;</TD>"
					End If
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("CT").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("BirthState").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("GenderShortName").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("JoinDate").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("CotDate").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Period").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("MaritalStatusID").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Address").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Colony").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("city").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ZipZone").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("State").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Nombram").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Afore").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ICEFA").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ICNumber").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Salary").Value), 2, True, False, True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Salary_v").Value), 2, True, False, True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("FullPay").Value), 2, True, False, True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("WorkingDays").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeContributions").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeContributionsAmount").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</TD>"
				sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do				
			Loop
			sRowContents = sRowContents & "</TABLE>"
			
			sQuery = "Select * From Dm_Deleted_HistoryList Where EmployeeID In (Select EmployeeID From Dm_Padron_Banamex Where Status=2)"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1000bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

			If Not oRecordset.EOF Then
				sDocumentName = sFilePath & "Rep_Referencias" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"

				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TR>"
					sRowContents = sRowContents & "<TD>Empresa</TD>"
					sRowContents = sRowContents & "<TD>No. Empleado</TD>"
					sRowContents = sRowContents & "<TD>RFC</TD>"
					sRowContents = sRowContents & "<TD>CURP</TD>"
					sRowContents = sRowContents & "<TD>Núm. Seguro Social</TD>"
					sRowContents = sRowContents & "<TD>Apellido paterno</TD>"
					sRowContents = sRowContents & "<TD>apellido materno</TD>"
					sRowContents = sRowContents & "<TD>Nombre</TD>"
					sRowContents = sRowContents & "<TD>CT</TD>"
					sRowContents = sRowContents & "<TD>Fecha de nacimiento</TD>"
					sRowContents = sRowContents & "<TD>Estado de nacimiento</TD>"
					sRowContents = sRowContents & "<TD>Género</TD>"
					sRowContents = sRowContents & "<TD>Fecha de contratación</TD>"
					sRowContents = sRowContents & "<TD>Fecha de cotización</TD>"
					sRowContents = sRowContents & "<TD>Periodo</TD>"
					sRowContents = sRowContents & "<TD>Estado civil</TD>"
					sRowContents = sRowContents & "<TD>Dirección</TD>"
					sRowContents = sRowContents & "<TD>Colonia</TD>"
					sRowContents = sRowContents & "<TD>Ciudad</TD>"
					sRowContents = sRowContents & "<TD>Código postal</TD>"
					sRowContents = sRowContents & "<TD>Estado</TD>"
					sRowContents = sRowContents & "<TD>Nombramiento</TD>"
					sRowContents = sRowContents & "<TD>Afore</TD>"
					sRowContents = sRowContents & "<TD>ICEFA</TD>"
					sRowContents = sRowContents & "<TD>Núm. Control Interno</TD>"
					sRowContents = sRowContents & "<TD>Salario</TD>"
					sRowContents = sRowContents & "<TD>Salario_v</TD>"
					sRowContents = sRowContents & "<TD>Salario integrado</TD>"
					sRowContents = sRowContents & "<TD>Días laborados</TD>"
					sRowContents = sRowContents & "<TD>Contribución del empleado</TD>"
					sRowContents = sRowContents & "<TD>Monto</TD>"
					sRowContents = sRowContents & "<TD>Fecha de inicio</TD>"
					sRowContents = sRowContents & "<TD>Fecha fin</TD>"
				sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("u_version").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("rfc").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("curp").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</TD>"
						Else
							sRowContents = sRowContents & "<TD>&nbsp;</TD>"
						End If
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("CT").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("BirthState").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("GenderShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("JoinDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("CotDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Period").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("MaritalStatusID").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Address").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Colony").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("city").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ZipZone").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("State").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Nombram").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("Afore").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ICEFA").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("ICNumber").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Salary").Value), 2, True, False, True) & "</TD>"
						sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Salary_v").Value), 2, True, False, True) & "</TD>"
						sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("FullPay").Value), 2, True, False, True) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("WorkingDays").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeContributions").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeContributionsAmount").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do				
				Loop
				sRowContents = sRowContents & "</TABLE>"
			End If

			If lErrorNumber = 0 Then
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
			End If
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
		Else
			lErrorNumber = -1
			sErrorDescription = "No se han detectado altas de personal para este periodo."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1035 = lErrorNumber
	Err.Clear
End Function
%>