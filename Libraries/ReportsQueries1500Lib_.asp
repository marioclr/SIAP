<%
Function BuildReport1502(oRequest, oADODBConnection, lEmployeeTypeID, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the amounts needed to pay the given positions
'Inputs:  oRequest, oADODBConnection, lEmployeeTypeID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1502"
	Dim sCondition
	Dim oRecordset
	Dim asAmounts
	Dim asPayrolls
	Dim dTotal
	Dim iIndex
	Dim jIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = sCondition & " And (Employees.EmployeeTypeID In (" & lEmployeeTypeID & "))"
	sCondition = Replace(sCondition, "EmployeeTypes.", "Employees.")
	sErrorDescription = "No se pudieron obtener los puestos por centro de trabajo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.LevelID, EconomicZoneID, Positions.PositionID, PositionShortName, PositionName, Count(EmployeeID) As PositionCounter From Employees, Jobs, Positions, Areas, Zones Where (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition & " Group By Employees.LevelID, EconomicZoneID, Positions.PositionID, PositionShortName, PositionName Order By Employees.LevelID, EconomicZoneID, PositionShortName, PositionName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Nivel,Zona<BR />Econ.,<SPAN COLS=""2"" />Puesto o Categoría,Número de<BR />plazas,Sueldo<BR />unitario,Compenzación<BR />Garantizada,Asignación<BR />Médica<BR />Unitaria,Ayuda para<BR />Gastos de<BR />Actualizacion<BR />Unitaria,SUELDO<BR />MENSUAL,SUELDO<BR />1103-001-00,COMPENSACIÓN<BR />GARANTIZADA,AJUSTE A<BR />CALENDARIO<BR />1103-002-00,P.VACAC.<BR />1305-018-00,G. FIN DE AÑO<BR />1306-026-00,ISR G. FIN<BR />DE AÑO,G. FIN DE AÑO<BR />DE LA<BR />COMPENSACIÓN,ISR G. FIN DE<BR />AÑO DE LA<BR />COMPENSACIÓN,AYUDA GASTOS<BR />DE ACTUAL.<BR />1325-048-00,APORTACIONES<BR />AL SEGURO DE<BR />SALUD<BR />1401-001-00,APOR. AL<BR />FOVISSSTE<BR />1403-000-00,CUOTAS SEG.<BR />PERS.CIVIL<BR />1404-000-00,SEG DE SEP.<BR />INDV 1407-000-00,ISR DE SEG<BR />DE SEP. INDV.,SEG. COL. DE<BR />RETIRO<BR />1408-000-00,APORT. AL<BR />SEG. DE<BR />RETIRO<BR />1414-001-00,APORT. AL SEG.<BR />DE CESANTIA EN<BR />EDAD AVAN. Y<BR />VEJEZ 1414-002-00,DEPOSITOS<BR />PARA<BR />AHORRO<BR />SOLIDARIO<BR />1415-000-00,C. P. F. DE<BR />AHORRO<BR />DEL P. C.<BR />1501-000-00,DESPENSA<BR />1507-010-00,BONO DE<BR />REYES<BR />1507-032-00,AYUDA DE<BR />TRANSP.<BR />1507-033-00,AYUD.<BR />COMPRA<BR />DE UTILES<BR />1507-034-00,VALE DE<BR />DESPENSA<BR />1512-047-00,PREV. DE<BR />AYUDA MUL.<BR />1511-002-00,ASIG. A<BR />PERS. MED.<BR />1512-035-00,PREMIO DE<BR />ANIVERSARIO<BR />1702-021-00,PREMIO 10 DE<BR />MAYO<BR />1702-022-00,DIAS ECON.<BR />NO DISFRUT.<BR />1702-024-00,ESTÍMULOS DE<BR />ASISTEN.<BR />1702-037-00,ESTÍMULOS DE<BR />PUNTUAL.<BR />1702-038-00,ESTÍMULOS DE<BR />DESEM.<BR />1702-039-00,EST. MERITO<BR />RELEVANTE<BR />1702-040-00,TOTAL", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split("CENTER,CENTER,CENTER,,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				asAmounts = Split(",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					asAmounts(0) = asAmounts(0) & Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000")) & LIST_SEPARATOR
					asAmounts(1) = asAmounts(1) & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value)) & LIST_SEPARATOR
					asAmounts(2) = asAmounts(2) & CStr(oRecordset.Fields("PositionID").Value) & SECOND_LIST_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & LIST_SEPARATOR
					asAmounts(3) = asAmounts(3) & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & LIST_SEPARATOR
					asAmounts(4) = asAmounts(4) & FormatNumber(CLng(oRecordset.Fields("PositionCounter").Value), 0, True, False, True) & LIST_SEPARATOR
					For iIndex = 5 To UBound(asAmounts)
						asAmounts(iIndex) = asAmounts(iIndex) & "0" & LIST_SEPARATOR
					Next
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				For iIndex = 0 To UBound(asAmounts)
					asAmounts(iIndex) = Left(asAmounts(iIndex), (Len(asAmounts(iIndex)) - Len(LIST_SEPARATOR)))
					asAmounts(iIndex) = Split(asAmounts(iIndex), LIST_SEPARATOR)
				Next
				For jIndex = 0 To UBound(asAmounts(2))
					asAmounts(2)(jIndex) = Split(asAmounts(2)(jIndex), SECOND_LIST_SEPARATOR)
				Next
				For jIndex = 0 To UBound(asAmounts(4))
					asAmounts(4)(jIndex) = CLng(asAmounts(4)(jIndex))
				Next
				For iIndex = 5 To UBound(asAmounts)
					For jIndex = 0 To UBound(asAmounts(iIndex))
						asAmounts(iIndex)(jIndex) = 0
					Next
				Next

'				sErrorDescription = "No se pudieron obtener los identificadores de las nóminas."
'				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollDate>=" & Year(Date()) - 1 & Right(("0" & Month(Date())), Len("00")) & "00) And (PayrollDate<=" & Year(Date()) & Right(("0" & Month(Date())), Len("00")) & "99) Order By PayrollID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'				If lErrorNumber = 0 Then
'					asPayrolls = ""
'					Do While Not oRecordset.EOF
'						asPayrolls = asPayrolls & CStr(oRecordset.Fields("PayrollID").Value) & ","
'						oRecordset.MoveNext
'						If Err.number <> 0 Then Exit Do
'					Loop
'					oRecordset.Close
'					If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))
'					asPayrolls = Split(asPayrolls, ",")
					For iIndex = 0 To UBound(asAmounts(0))
'						For jIndex = 0 To UBound(asPayrolls)
'							sErrorDescription = "No se pudieron obtener los montos pagados en las nóminas, agrupados por concepto."
'							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept01 From Payroll_" & asPayrolls(jIndex) & ", Employees, Jobs, Positions, Areas Where (Payroll_" & asPayrolls(jIndex) & ".EmployeeID=Employees.EmployeeID) And (ConceptID=1) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Employees.LevelID=" & asAmounts(0)(iIndex) & ") And (Areas.EconomicZoneID=" & asAmounts(1)(iIndex) & ") And (Positions.PositionID=" & asAmounts(2)(iIndex)(0) & ")", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							sErrorDescription = "No se pudieron obtener los montos del tabulador."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept01 From ConceptsValues Where (ConceptID=1) And (LevelID In (-1," & asAmounts(0)(iIndex) & ")) And (EconomicZoneID In (0," & asAmounts(1)(iIndex) & ")) And (PositionID <> -1)" & " And (EmployeeTypeID In (" & lEmployeeTypeID & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									If asAmounts(5)(iIndex) < CDbl(oRecordset.Fields("MaxConcept01").Value) Then asAmounts(5)(iIndex) = CDbl(oRecordset.Fields("MaxConcept01").Value)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
							End If
							asAmounts(5)(iIndex) = asAmounts(5)(iIndex) * 2

'							sErrorDescription = "No se pudieron obtener los montos pagados en las nóminas, agrupados por concepto."
'							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept03 From Payroll_" & asPayrolls(jIndex) & ", Employees, Jobs, Positions, Areas Where (Payroll_" & asPayrolls(jIndex) & ".EmployeeID=Employees.EmployeeID) And (ConceptID=3) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Employees.LevelID=" & asAmounts(0)(iIndex) & ") And (Areas.EconomicZoneID=" & asAmounts(1)(iIndex) & ") And (Positions.PositionID=" & asAmounts(2)(iIndex)(0) & ")", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							sErrorDescription = "No se pudieron obtener los montos del tabulador."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept03 From ConceptsValues Where (ConceptID=3) And (LevelID In (-1," & asAmounts(0)(iIndex) & ")) And (EconomicZoneID In (0," & asAmounts(1)(iIndex) & ")) And (PositionID=" & asAmounts(2)(iIndex)(0) & ") And (EmployeeTypeID In (" & lEmployeeTypeID & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									If asAmounts(6)(iIndex) < CDbl(oRecordset.Fields("MaxConcept03").Value) Then asAmounts(6)(iIndex) = CDbl(oRecordset.Fields("MaxConcept03").Value)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
							End If
							asAmounts(6)(iIndex) = asAmounts(6)(iIndex) * 2

'							sErrorDescription = "No se pudieron obtener los montos pagados en las nóminas, agrupados por concepto."
'							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept35 From Payroll_" & asPayrolls(jIndex) & ", Employees, Jobs, Positions, Areas Where (Payroll_" & asPayrolls(jIndex) & ".EmployeeID=Employees.EmployeeID) And (ConceptID=38) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Employees.LevelID=" & asAmounts(0)(iIndex) & ") And (Areas.EconomicZoneID=" & asAmounts(1)(iIndex) & ") And (Positions.PositionID=" & asAmounts(2)(iIndex)(0) & ")", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							sErrorDescription = "No se pudieron obtener los montos del tabulador."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept35 From ConceptsValues Where (ConceptID=38) And (LevelID In (-1," & asAmounts(0)(iIndex) & ")) And (EconomicZoneID In (0," & asAmounts(1)(iIndex) & ")) And (PositionID=" & asAmounts(2)(iIndex)(0) & ") And (EmployeeTypeID In (" & lEmployeeTypeID & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									If asAmounts(7)(iIndex) < CDbl(oRecordset.Fields("MaxConcept35").Value) Then asAmounts(7)(iIndex) = CDbl(oRecordset.Fields("MaxConcept35").Value)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
							End If
							asAmounts(7)(iIndex) = asAmounts(7)(iIndex) * 2

'							sErrorDescription = "No se pudieron obtener los montos pagados en las nóminas, agrupados por concepto."
'							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept48 From Payroll_" & asPayrolls(jIndex) & ", Employees, Jobs, Positions, Areas Where (Payroll_" & asPayrolls(jIndex) & ".EmployeeID=Employees.EmployeeID) And (ConceptID=49) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Employees.LevelID=" & asAmounts(0)(iIndex) & ") And (Areas.EconomicZoneID=" & asAmounts(1)(iIndex) & ") And (Positions.PositionID=" & asAmounts(2)(iIndex)(0) & ")", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							sErrorDescription = "No se pudieron obtener los montos del tabulador."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept48 From ConceptsValues Where (ConceptID=49) And (LevelID In (-1," & asAmounts(0)(iIndex) & ")) And (EconomicZoneID In (0," & asAmounts(1)(iIndex) & ")) And (PositionID=" & asAmounts(2)(iIndex)(0) & ") And (EmployeeTypeID In (" & lEmployeeTypeID & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									If asAmounts(8)(iIndex) < CDbl(oRecordset.Fields("MaxConcept48").Value) Then asAmounts(8)(iIndex) = CDbl(oRecordset.Fields("MaxConcept48").Value)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
							End If
							asAmounts(8)(iIndex) = asAmounts(8)(iIndex) * 2
'						Next

						asAmounts(9)(iIndex) = asAmounts(5)(iIndex) + asAmounts(6)(iIndex) + asAmounts(7)(iIndex) + asAmounts(8)(iIndex)
						asAmounts(10)(iIndex) = asAmounts(5)(iIndex) * asAmounts(4)(iIndex) * 12
						asAmounts(11)(iIndex) = asAmounts(6)(iIndex) * asAmounts(4)(iIndex) * 12
						asAmounts(12)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(11)(iIndex)) / 30 * 5 * asAmounts(4)(iIndex)
						asAmounts(13)(iIndex) = (asAmounts(5)(iIndex) / 30) * 10 * asAmounts(4)(iIndex)
						asAmounts(14)(iIndex) = (asAmounts(5)(iIndex) / 30) * 40 * asAmounts(4)(iIndex)
						asAmounts(15)(iIndex) = asAmounts(15)(iIndex) * 0.16
						asAmounts(16)(iIndex) = (asAmounts(6)(iIndex) / 30) * 40 * asAmounts(4)(iIndex)
						asAmounts(17)(iIndex) = asAmounts(16)(iIndex) * 0.20
						asAmounts(18)(iIndex) = asAmounts(8)(iIndex) * 12 * asAmounts(4)(iIndex)
						asAmounts(19)(iIndex) = (asAmounts(10)(iIndex) * 0.0997) + 2682.12 * asAmounts(4)(iIndex)
						asAmounts(20)(iIndex) = (asAmounts(5)(iIndex) * 0.05) * 12 * asAmounts(4)(iIndex)
						asAmounts(21)(iIndex) = (asAmounts(5)(iIndex) * 0.0229) * 12 * asAmounts(4)(iIndex)
						asAmounts(22)(iIndex) = 0
						asAmounts(23)(iIndex) = 0
						asAmounts(24)(iIndex) = asAmounts(4)(iIndex) * 13.49 * 12
						asAmounts(25)(iIndex) = (asAmounts(5)(iIndex) * 0.02) * 12 * asAmounts(4)(iIndex)
						asAmounts(26)(iIndex) = (asAmounts(5)(iIndex) * 0.03175) + 1061.28 * asAmounts(4)(iIndex)
						asAmounts(27)(iIndex) = (asAmounts(5)(iIndex) * 0.065) * 12 * asAmounts(4)(iIndex)
						asAmounts(28)(iIndex) = (201.29 * 2) * 12 * asAmounts(4)(iIndex)
						asAmounts(29)(iIndex) = 77 * 12 * asAmounts(4)(iIndex)
						asAmounts(30)(iIndex) = 935.26 * asAmounts(4)(iIndex)
						asAmounts(31)(iIndex) = asAmounts(4)(iIndex) * 80 * 12
						asAmounts(32)(iIndex) = 110 * asAmounts(4)(iIndex)
						asAmounts(33)(iIndex) = 0
						asAmounts(34)(iIndex) = 95.58 * 12 * asAmounts(4)(iIndex)
						asAmounts(35)(iIndex) = asAmounts(7)(iIndex) * 12 * asAmounts(4)(iIndex)
						asAmounts(36)(iIndex) = 990.27 * asAmounts(4)(iIndex)
						asAmounts(37)(iIndex) = 1320.36 * asAmounts(4)(iIndex) / 2
						asAmounts(38)(iIndex) = 0
						asAmounts(39)(iIndex) = 0
						asAmounts(40)(iIndex) = 0
						asAmounts(41)(iIndex) = 0
						asAmounts(42)(iIndex) = 0
						asAmounts(43)(iIndex) = 0
					Next
'				End If

				For jIndex = 0 To UBound(asAmounts(0))
					sRowContents = ""
					dTotal = 0
					For iIndex = 0 To UBound(asAmounts)
						If iIndex = 2 Then
							sRowContents = sRowContents & asAmounts(2)(jIndex)(1) & TABLE_SEPARATOR
						ElseIf (iIndex > 4) And (iIndex < 43) Then
							sRowContents = sRowContents & FormatNumber(asAmounts(iIndex)(jIndex), 2, True, False, True) & TABLE_SEPARATOR
							If (iIndex > 9) And (iIndex < 43) Then dTotal = dTotal + asAmounts(iIndex)(jIndex)
						ElseIf iIndex = 43 Then
							asAmounts(43)(jIndex) = dTotal
							sRowContents = sRowContents & FormatNumber(dTotal, 2, True, False, True) & TABLE_SEPARATOR
						Else
							sRowContents = sRowContents & asAmounts(iIndex)(jIndex) & TABLE_SEPARATOR
						End If
					Next
					sRowContents = Left(sRowContents, (Len(sRowContents) - Len(TABLE_SEPARATOR)))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				Next

				sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
				dTotal = 0
				For iIndex = 0 To UBound(asAmounts(4))
					dTotal = dTotal + asAmounts(4)(iIndex)
				Next
				sRowContents = sRowContents & FormatNumber(dTotal, 0, True, False, True) & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
				For iIndex = 10 To UBound(asAmounts)
					dTotal = 0
					For jIndex = 0 To UBound(asAmounts(0))
						dTotal = dTotal + asAmounts(iIndex)(jIndex)
					Next
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dTotal, 2, True, False, True)
				Next
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
'		Else
'			lErrorNumber = L_ERR_NO_RECORDS
'			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1502 = lErrorNumber
	Err.Clear
End Function

Function BuildReports1502(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To call BuildReport1332 for each Employee Type
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReports1502"
	Dim asEmployeeTypes
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudieron obtener los tipos de empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypeID, EmployeeTypeName From EmployeeTypes Where (EmployeeTypeID>-1) And (Active=1) Order By EmployeeTypeName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		asEmployeeTypes = ""
		Do While Not oRecordset.EOF
			asEmployeeTypes = asEmployeeTypes & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "," & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & LIST_SEPARATOR
			oRecordset.MoveNext
		Loop
		oRecordset.Close
		If Len(asEmployeeTypes) > 0 Then asEmployeeTypes = Left(asEmployeeTypes, (Len(asEmployeeTypes) - Len(LIST_SEPARATOR)))
		asEmployeeTypes = Split(asEmployeeTypes, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asEmployeeTypes)
			asEmployeeTypes(iIndex) = Split(asEmployeeTypes(iIndex), ",", 2)
			Response.Write "<B>" & asEmployeeTypes(iIndex)(1) & "</B><BR />"
			lErrorNumber = BuildReport1332(oRequest, oADODBConnection, asEmployeeTypes(iIndex)(0), bForExport, sErrorDescription)
			Response.Write "<BR /><BR />"
		Next
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReports1502 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1503(oRequest, oADODBConnection, lReportID, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the amounts needed to pay the given positions
'Inputs:  oRequest, oADODBConnection, lReportID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1503"
	Dim sCondition
	Dim oRecordset
	Dim asAmounts
	Dim asConditions
	Dim asParameters
	Dim dTotal
	Dim lPeriods
	Dim iIndex
	Dim jIndex
	Dim asColumnsTitles
	Dim asRowsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	asAmounts = Split(BuildList("", ",", 102), ",")
	asConditions = Split(",,,,,,,", ",", -1, vbBinaryCompare)
	asParameters = Split(",,,,,,", ",")
	For iIndex = 0 To UBound(asParameters)
		asParameters(iIndex) = Split(BuildList("", ",", 77), ",")
		For jIndex = 0 To UBound(asParameters(iIndex))
			asParameters(iIndex)(jIndex) = Split((oRequest("P_" & iIndex & "_" & jIndex & "_0").Item & LIST_SEPARATOR & oRequest("P_" & iIndex & "_" & jIndex & "_1").Item), LIST_SEPARATOR)
			asParameters(iIndex)(jIndex)(0) = CDbl(asParameters(iIndex)(jIndex)(0))
			asParameters(iIndex)(jIndex)(1) = CDbl(asParameters(iIndex)(jIndex)(1))
		Next
	Next
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(Replace(Replace(sCondition, "Employees", "BudgetsPositions"), "EmployeeTypes", "BudgetsPositions"), "EconomicZones", "BudgetsPositions"), "Companies", "BudgetsPositions")
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetsPositions.PositionID, PositionShortName, PositionName, BudgetsPositions.EmployeeTypeID, BudgetsPositions.CompanyID, BudgetsPositions.ClassificationID, BudgetsPositions.GroupGradeLevelID, BudgetsPositions.IntegrationID, BudgetsPositions.LevelID, BudgetsPositions.EconomicZoneID, TotalPositions, CompanyName, LevelName, GroupGradeLevelName, EmployeeTypeShortName, EmployeeTypeName From BudgetsPositions, Companies, Levels, GroupGradeLevels, EmployeeTypes Where (BudgetsPositions.CompanyID=Companies.CompanyID) And (BudgetsPositions.LevelID=Levels.LevelID) And (BudgetsPositions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (BudgetsPositions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (BudgetsPositions.EndDate=30000000) " & sCondition & " Order By PositionShortName, EmployeeTypeShortName, CompanyName, EconomicZoneID, LevelName, GroupGradeLevelName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				asAmounts(0) = asAmounts(0) & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
				asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
					asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value))
				Else
					asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value))
				End If
				asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
				asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
				asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value))
				If Len(oRequest("TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value)).Item) = 0 Then
					asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("TotalPositions").Value), 0, True, False, True)
				Else
					asAmounts(0) = asAmounts(0) & TABLE_SEPARATOR & FormatNumber(CLng(oRequest("TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value)).Item), 0, True, False, True)
				End If
				asAmounts(0) = asAmounts(0) & LIST_SEPARATOR & SECOND_LIST_SEPARATOR & LIST_SEPARATOR

				asAmounts(1) = asAmounts(1) & CStr(oRecordset.Fields("LevelID").Value) & LIST_SEPARATOR
				asAmounts(2) = asAmounts(2) & CStr(oRecordset.Fields("EconomicZoneID").Value) & LIST_SEPARATOR
				asAmounts(3) = asAmounts(3) & CStr(oRecordset.Fields("PositionID").Value) & LIST_SEPARATOR

				asConditions(0) = asConditions(0) & CStr(oRecordset.Fields("CompanyID").Value) & LIST_SEPARATOR
				asConditions(1) = asConditions(1) & CStr(oRecordset.Fields("EmployeeTypeID").Value) & LIST_SEPARATOR
				asConditions(2) = asConditions(2) & CStr(oRecordset.Fields("ClassificationID").Value) & LIST_SEPARATOR
				asConditions(3) = asConditions(3) & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & LIST_SEPARATOR
				asConditions(4) = asConditions(4) & CStr(oRecordset.Fields("IntegrationID").Value) & LIST_SEPARATOR
				asConditions(5) = asConditions(5) & CStr(oRecordset.Fields("LevelID").Value) & LIST_SEPARATOR
				asConditions(6) = asConditions(6) & CStr(oRecordset.Fields("EconomicZoneID").Value) & LIST_SEPARATOR
				asConditions(7) = asConditions(7) & CStr(oRecordset.Fields("PositionID").Value) & LIST_SEPARATOR

				If Len(oRequest("TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value)).Item) = 0 Then
					asAmounts(4) = asAmounts(4) & CStr(oRecordset.Fields("TotalPositions").Value) & LIST_SEPARATOR
				Else
					asAmounts(4) = asAmounts(4) & oRequest("TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value)).Item & LIST_SEPARATOR
				End If
				For iIndex = 5 To UBound(asAmounts)
					asAmounts(iIndex) = asAmounts(iIndex) & "0" & LIST_SEPARATOR
				Next
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asAmounts(0) = Left(asAmounts(0), (Len(asAmounts(0)) - Len(LIST_SEPARATOR & SECOND_LIST_SEPARATOR & LIST_SEPARATOR)))
			asAmounts(0) = Split(asAmounts(0), LIST_SEPARATOR & SECOND_LIST_SEPARATOR & LIST_SEPARATOR)
			For iIndex = 1 To UBound(asAmounts)
				asAmounts(iIndex) = Left(asAmounts(iIndex), (Len(asAmounts(iIndex)) - Len(LIST_SEPARATOR)))
				asAmounts(iIndex) = Split(asAmounts(iIndex), LIST_SEPARATOR)
			Next
			For iIndex = 0 To UBound(asConditions)
				asConditions(iIndex) = Left(asConditions(iIndex), (Len(asConditions(iIndex)) - Len(LIST_SEPARATOR)))
				asConditions(iIndex) = Split(asConditions(iIndex), LIST_SEPARATOR)
			Next
			For jIndex = 0 To UBound(asAmounts(4))
				asAmounts(4)(jIndex) = CInt(asAmounts(4)(jIndex))
			Next
			For iIndex = 5 To UBound(asAmounts)
				For jIndex = 0 To UBound(asAmounts(iIndex))
					asAmounts(iIndex)(jIndex) = 0
				Next
			Next

			lPeriods = asParameters(CInt(asConditions(1)(0)))(76)(0)
			For iIndex = 0 To UBound(asAmounts(1))
				sErrorDescription = "No se pudieron obtener los montos del Sueldo base."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept01 From ConceptsValues Where (ConceptID=1) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept01").Value) Then
							If asAmounts(5)(iIndex) < CDbl(oRecordset.Fields("MaxConcept01").Value) Then asAmounts(5)(iIndex) = CDbl(oRecordset.Fields("MaxConcept01").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(5)(iIndex) = asAmounts(5)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos de la Compensación garantizada."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept03 From ConceptsValues Where (ConceptID=3) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept03").Value) Then
							If asAmounts(6)(iIndex) < CDbl(oRecordset.Fields("MaxConcept03").Value) Then asAmounts(6)(iIndex) = CDbl(oRecordset.Fields("MaxConcept03").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(6)(iIndex) = asAmounts(6)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos de la Compensación por riesgos profesionales."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept04 From ConceptsValues Where (ConceptID=4) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept04").Value) Then
							If asAmounts(8)(iIndex) < CDbl(oRecordset.Fields("MaxConcept04").Value) Then asAmounts(8)(iIndex) = CDbl(oRecordset.Fields("MaxConcept04").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(7)(iIndex) = asAmounts(7)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos de la Asignación a personal médico y paramédicos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept35 From ConceptsValues Where (ConceptID=38) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept35").Value) Then
							If asAmounts(7)(iIndex) < CDbl(oRecordset.Fields("MaxConcept35").Value) Then asAmounts(7)(iIndex) = CDbl(oRecordset.Fields("MaxConcept35").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(8)(iIndex) = asAmounts(8)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos de la Ayuda gastos de actualización."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept48 From ConceptsValues Where (ConceptID=49) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept48").Value) Then
							If asAmounts(9)(iIndex) < CDbl(oRecordset.Fields("MaxConcept48").Value) Then asAmounts(9)(iIndex) = CDbl(oRecordset.Fields("MaxConcept48").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(9)(iIndex) = asAmounts(9)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos de la Becas médicos residentes."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConceptB2 From ConceptsValues Where (ConceptID=89) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConceptB2").Value) Then
							If asAmounts(10)(iIndex) < CDbl(oRecordset.Fields("MaxConceptB2").Value) Then asAmounts(10)(iIndex) = CDbl(oRecordset.Fields("MaxConceptB2").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(10)(iIndex) = asAmounts(10)(iIndex) * 2

				sErrorDescription = "No se pudieron obtener los montos del tabulador."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxConcept36 From ConceptsValues Where (ConceptID=39) And (EmployeeTypeID In (-1," & asConditions(1)(iIndex) & ")) And (ClassificationID In (-1," & asConditions(2)(iIndex) & ")) And (GroupGradeLevelID In (-1," & asConditions(3)(iIndex) & ")) And (IntegrationID In (-1," & asConditions(4)(iIndex) & ")) And (LevelID In (-1," & asConditions(5)(iIndex) & ")) And (EconomicZoneID In (0," & asConditions(6)(iIndex) & "))", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If Not IsNull(oRecordset.Fields("MaxConcept36").Value) Then
							If asAmounts(11)(iIndex) < CDbl(oRecordset.Fields("MaxConcept36").Value) Then asAmounts(11)(iIndex) = CDbl(oRecordset.Fields("MaxConcept36").Value)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				asAmounts(11)(iIndex) = asAmounts(11)(iIndex) * 2

				asAmounts(12)(iIndex) = asAmounts(5)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Sueldos base = Sueldo base * Plazas * Periodo * Previsión incremento
				asAmounts(13)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Honorarios = xxx * Plazas * Periodo * Previsión incremento
				asAmounts(14)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Sueldo base al personal eventual = xxx * Plazas * Periodo * Previsión incremento
				asAmounts(15)(iIndex) = asAmounts(10)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Beca a médicos residentes = Beca a médicos residentes * Plazas * Periodo * Previsión incremento
				asAmounts(16)(iIndex) = asAmounts(11)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Complemento de beca = Complemento de beca * Plazas * Periodo * Previsión incremento
				asAmounts(17)(iIndex) = asAmounts(6)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Compensación garantizada = Compensación garantizada * Plazas * Periodo * Previsión incremento
				asAmounts(18)(iIndex) = asAmounts(8)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'Asignación médica = Asignación médica * Plazas * Periodo * Previsión incremento
				asAmounts(19)(iIndex) = asAmounts(9)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * ((100 + asParameters(CInt(asConditions(1)(iIndex)))(74)(0)) / 100) 'AGA = Ayuda de gastos de actualización * Plazas * Periodo * Previsión incremento
				asAmounts(23)(iIndex) = asAmounts(5)(iIndex) / 30 * 10 * asAmounts(4)(iIndex) 'Prima vacacional = Sueldo base / 30 * 10 * Plazas
				asAmounts(24)(iIndex) = asAmounts(5)(iIndex) / 30 * 40 * (asParameters(CInt(asConditions(1)(iIndex)))(76)(0) /12) * asAmounts(4)(iIndex) 'Aguinaldo = Sueldo base / 30 * 40 * (Periodo / 12) * Plazas
				asAmounts(27)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(3)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Despensa = Despensa * Plazas * Periodo
				asAmounts(28)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(54)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al Sistema de Ahorro para el Retiro (1508) = Sueldo base * Factor SAR * Plazas * Periodo
				asAmounts(30)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(49)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al ISSSTE = Sueldo base * Factor ISSSTE * Plazas * Periodo
				asAmounts(31)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(48)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuota social = Cuota ISSSTE * Plazas * Periodo
				asAmounts(32)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(50)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al FOVISSSTE = Sueldo base * Factor FOVISSSTE * Plazas * Periodo
				asAmounts(33)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) * 0.0229 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Seguro institucional = (Sueldo base + compensación garantizada) * 2.29% * Plazas * Periodo
				asAmounts(34)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) * asParameters(CInt(asConditions(1)(iIndex)))(65)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el seguro de separación individualizado = (Sueldo base + Compensación garantizada) * Factor SSI * Plazas * Periodo
				asAmounts(35)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(57)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el seguro colectivo de retiro = Seguro colectivo * Plazas * Periodo
				asAmounts(38)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(54)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al Sistema de Ahorro para el Retiro (1413) = Sueldo base * Factor SAR * Plazas * Periodo
				asAmounts(39)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(39)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(39)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Prima quinquenal por años de servicios efectivos prestados = Prima quinquenal * (Plazas * Factor de plazas) * Periodo
				asAmounts(43)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) * asParameters(CInt(asConditions(1)(iIndex)))(66)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el seguro de gastos médicos del personal civil = (Sueldo base + compensación garantizada) * Factor NSI * Plazas * Periodo
				asAmounts(44)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(58)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Seguro responsabilidad civil = Seguro responsabilidad civil * Plazas * Periodo
				asAmounts(46)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(56)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al seguro por cesantía = Sueldo base * Factor cesantía * Plazas * Periodo
				asAmounts(47)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(55)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuota de cesantía = Cuota cesantía * Plazas * Periodo
				asAmounts(48)(iIndex) = asAmounts(5)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(51)(0) / 100 * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Depósitos para el ahorro solidario = Sueldo base * Factor ahorro solidario * Plazas * Periodo
				asAmounts(55)(iIndex) = asAmounts(6)(iIndex) / 30 * 40 * asAmounts(4)(iIndex) 'Gratificación de fin de año de la compensación garantizada = Compensación garantizada / 30 * 40 * Plazas
				asAmounts(96)(iIndex) = asAmounts(24)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(72)(0) / 100 'ISR gratificación de fin de año = Aguinaldo * Factor de ISR (aguinaldo)
				asAmounts(97)(iIndex) = asAmounts(55)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(73)(0) / 100 'ISR Gratificación de fin de año de la compensación garantizada = Gratificación de fin de año de la compensación garantizada * Factor de ISR (aguinaldo de la compensación)
				asAmounts(98)(iIndex) = asAmounts(34)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(71)(0) / 100 'ISR seguro de separación individualizado = Cuotas para el seguro de separación individualizado * Factor de ISR (SSI)

				Select Case CInt(asConditions(1)(iIndex))
					Case 0, 5 'Médica, paramédica y grupos afines
						asAmounts(20)(iIndex) = asAmounts(5)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * asParameters(CInt(asConditions(1)(iIndex)))(46)(0) / 100 'Subsidio para el empleo = Sueldo base * Plazas * Periodo * Subsidio para el empleo
						asAmounts(21)(iIndex) = asAmounts(15)(iIndex) * asAmounts(4)(iIndex) 'Gratificación de becas = Beca a médicos residentes * Plazas
						asAmounts(22)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(60)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(60)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Retribuciones por servicios de carácter social = Retribuciones por servicios de carácter social * (Plazas * Factor de plazas) * Periodo
						asAmounts(23)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(15)(iIndex)) / 30 * 10 * asAmounts(4)(iIndex) 'Prima vacacional = (Sueldo base + Beca a médicos residentes) / 30 * 10 * Plazas
						asAmounts(24)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(15)(iIndex)) / 30 * 40 * (asParameters(CInt(asConditions(1)(iIndex)))(76)(0) /12) * asAmounts(4)(iIndex) 'Aguinaldo = (Sueldo base + Beca a médicos residentes) / 30 * 40 * (Periodo / 12) * Plazas
						asAmounts(25)(iIndex) = (asAmounts(5)(iIndex) * 0.2) + (asAmounts(5)(iIndex) * 0.1) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por riesgos profesionales = (Sueldo base * 20%) + (Sueldo base * 10%) * Plazas * Periodo
						asAmounts(26)(iIndex) = (asAmounts(5)(iIndex) * 0.2) + (asAmounts(5)(iIndex) * 0.1) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación a médicos residentes = (Sueldo base * 20%) + (Sueldo base * 10%) * Plazas * Periodo
						asAmounts(29)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(6)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Previsión social múltiple = Previsión social múltiple * Plazas * Periodo
						asAmounts(41)(iIndex) = asAmounts(5)(iIndex) / 30 / 8 * 36 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(7)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Horas extras = Sueldo / 30 / 8 * 36 * (Plazas * Factor de plazas) * Periodo
						asAmounts(49)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 3 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el fondo de ahorro del personal civil = SMB * 3 * Plazas * Periodo
						asAmounts(68)(iIndex) = asAmounts(5)(iIndex) / 6.5 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(11)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Turno opcional = Sueldo base / 6.5 * 3 * (Plazas * Factor de Plazas) * Periodo
						asAmounts(54)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex) + asAmounts(18)(iIndex) + asAmounts(19)(iIndex) + asAmounts(26)(iIndex)) / 30 * asParameters(CInt(asConditions(1)(iIndex)))(5)(0) * asAmounts(4)(iIndex) 'Ajuste al calendario = (Sueldo + turno opcional + percepción adicional + asignación a personal médico + ayuda de gastos de actualización + compensacion a médicos residentes) / 30 * Ajuste al calendario * Plazas
						asAmounts(56)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(64)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(64)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda renta a becarios = Ayuda renta becarios * (Plazas * Factor plazas) * Periodo
						asAmounts(57)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(15)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda de transporte = Ayuda de transporte * Plazas * Periodo
						asAmounts(58)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(4)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Apoyo para el desarrollo y capacitación = CODECA * Plazas * Periodo
						asAmounts(59)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(16)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Vales de despensa = Vales de despensa * Plazas * Periodo
						asAmounts(60)(iIndex) = asAmounts(15)(iIndex) * 4 * asAmounts(4)(iIndex) 'Material didáctico = Beca a médicos residentes * 4 * Plazas
						asAmounts(61)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 9 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(42)(1) / 100) + 0.9999999999999) 'Premio de aniversario = SMB * 9 * (Plazas * Factor de plazas)
						asAmounts(62)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(40)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por antigüedad = Compensación por antigüedad * (Plazas * Factor de plazas) * Periodo
						asAmounts(63)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 0.25 * 48 * asAmounts(4)(iIndex) 'Prima dominical = (Sueldo + turno opcional + percepción adicional) / 30 * 25% * 48 * Plazas
						asAmounts(64)(iIndex) = asAmounts(5)(iIndex) / 30 * 7 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(10)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Remuneraciones por suplencias = Sueldo base / 30 * 7 * (Plazas * Factor de plazas) * Periodo
						asAmounts(65)(iIndex) = asAmounts(5)(iIndex) / 30 / 8 * 2 * 60 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(9)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Remuneraciones por guardias = Sueldo / 30 / 8 * 2 * 60 * (Plazas * Factor de plazas) * Periodo
						asAmounts(66)(iIndex) = asAmounts(5)(iIndex) / 30 / 8 * 2 * 60 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(61)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Servicios prioritarios de atención primaria a la Salud = Sueldo / 30 / 8 * 2 * 60 * (Plazas * Factor de plazas) * Periodo
						asAmounts(67)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(63)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Rezago quirúrgico = Rezago quirúrgico * Plazas * Periodo
						asAmounts(69)(iIndex) = asAmounts(5)(iIndex) / 6.5 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(8)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Remuneración adicional = Sueldo base / 6.5 * 3 * (Plazas * Factor de plazas) * Periodo
						asAmounts(70)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(62)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras = Cuota para el seguro * Plazas * Periodo
						asAmounts(71)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 8.5 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(17)(1) / 100) + 0.9999999999999) 'Bono de Reyes = SMB * 8.5 * (Plazas * Factor de plazas)
						asAmounts(72)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(18)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(18)(1) / 100) + 0.9999999999999) 'Ayuda compra de útiles = Ayuda compra de útiles * (Plazas * Factor de plazas)
						asAmounts(73)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(20)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(20)(1) / 100) + 0.9999999999999) 'Ayuda de anteojos = Ayuda de anteojos * (Plazas * Factor de plazas)
						asAmounts(74)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(22)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(22)(1) / 100) + 0.9999999999999) 'Ayuda por muerte familiar 1° grado = Ayuda por muerte familiar 1º grado * (Plazas * Factor de plazas)
						asAmounts(75)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(21)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(21)(1) / 100) + 0.9999999999999) 'Impresión de tesis = Ayuda para impresión de tesis * (Plazas * Factor de plazas)
						asAmounts(76)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(23)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(23)(1) / 100) + 0.9999999999999) 'Evento 10 de mayo = Evento 10 de mayo * (Plazas * Factor de plazas)
						asAmounts(77)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(24)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(24)(1) / 100) + 0.9999999999999) 'Evento día del niño = 120 * (Plazas * Factor de plazas)
						asAmounts(78)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 8 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(26)(1) / 100) + 0.9999999999999) 'Evento fomento cultural, turístico y deportivo = SMB * 8 * (Plazas * Factor de plazas)
						asAmounts(79)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(25)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(25)(1) / 100) + 0.9999999999999) 'Evento día del trabajador = Sueldo base * Evento día del trabajador * (Plazas * Factor de Plazas)
						asAmounts(80)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(29)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(29)(1) / 100) + 0.9999999999999) 'Comisión nacional de auxilio = Comisión nacional de auxilio * (Plazas * Factor de plazas)
						asAmounts(81)(iIndex) = asAmounts(5)(iIndex) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(14)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Jornada nocturna adicional por día festivo = Sueldo base * 30 * (Plazas * Factor de plazas) * Periodo
						asAmounts(82)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(31)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(31)(1) / 100) + 0.9999999999999)  'Premios, estímulos y recompensas = Premio * (Plazas * Factor de plazas)
						asAmounts(83)(iIndex) = asAmounts(5)(iIndex) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(27)(1) / 100) + 0.9999999999999) 'Participación de inventarios físicos = Sueldo base * (Plazas * Factor de plazas)
						asAmounts(84)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(28)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(28)(1) / 100) + 0.9999999999999) 'Becas hijos de trabajadores = Becas hijos de trabajadores * (Plazas * Factor de plazas)
						asAmounts(85)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 12 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(32)(1) / 100) + 0.9999999999999) 'Premio 10 de mayo = SMB * 12 * (Plazas * Factor de plazas)
						asAmounts(86)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) 'Premio por antigüedad = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas)
						asAmounts(87)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(12)(1) / 100) + 0.9999999999999) 'Días económicos no disfrutados = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas)
						asAmounts(89)(iIndex) = asAmounts(5)(iIndex) * 0.4 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería = Sueldo * 40% * Plazas * Periodo
						asAmounts(90)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(34)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por asistencia = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas) * Periodo
						asAmounts(91)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(36)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por puntualidad = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas) * Periodo
						asAmounts(92)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 2 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(36)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por desempeño = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * 2 días * (Plazas * Factor de plazas) * Periodo
						asAmounts(93)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex) + asAmounts(39)(iIndex)) / 30 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(37)(1) / 100) + 0.9999999999999) * 4 'Estímulo mérito relevante = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional + prima quinquenal) / 30 * 3 * (Plazas * Factor de plazas) * 4
						asAmounts(94)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(44)(1) / 100) + 0.9999999999999) 'Premio de antigüedad 25 y 30 años = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) * (Plazas * Factor de plazas)
						asAmounts(99)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(33)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(33)(1) / 100) + 0.9999999999999) 'Aportaciones por servicios de atención para el bienestar y desarrollo infantil = Cuota guardería * plazas * factor sobre plazas
					Case 2 'Operativos
						asAmounts(20)(iIndex) = asAmounts(5)(iIndex) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) * asParameters(CInt(asConditions(1)(iIndex)))(46)(0) / 100 'Subsidio para el empleo = Sueldo base * Plazas * Periodo * Subsidio para el empleo
						asAmounts(25)(iIndex) = (asAmounts(5)(iIndex) * 0.2) + (asAmounts(5)(iIndex) * 0.1) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por riesgos profesionales = (Sueldo base * 20%) + (Sueldo base * 10%) * Plazas * Periodo
						asAmounts(29)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(6)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Previsión social múltiple = Previsión social múltiple * Plazas * Periodo
						asAmounts(41)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) / 30 / 8 * 36 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(7)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Horas extras = Sueldo / 30 / 8 * 36 * (Plazas * Factor de plazas) * Periodo
						asAmounts(49)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 3 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el fondo de ahorro del personal civil = SMB * 3 * Plazas * Periodo
						asAmounts(54)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) / 30 * asParameters(CInt(asConditions(1)(iIndex)))(5)(0) * asAmounts(4)(iIndex) 'Ajuste al calendario = (Sueldo + Compensación garantizada) / 30 * Ajuste al calendario * No. Plazas
						asAmounts(57)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(15)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda de transporte = Ayuda de transporte * Plazas * Periodo
						asAmounts(58)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(4)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Apoyo para el desarrollo y capacitación = CODECA * Plazas * Periodo
						asAmounts(59)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(16)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Vales de despensa = Vales de despensa * Plazas * Periodo
						asAmounts(61)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 9 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(42)(1) / 100) + 0.9999999999999) 'Premio de aniversario = SMB * 9 * (Plazas * Factor de plazas)
						asAmounts(62)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(40)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por antigüedad = Compensación por antigüedad * (Plazas * Factor de plazas) * Periodo
						asAmounts(68)(iIndex) = asAmounts(5)(iIndex) / 6.5 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(11)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Turno opcional = Sueldo base / 6.5 * 3 * (Plazas * Factor de Plazas) * Periodo
						asAmounts(69)(iIndex) = asAmounts(5)(iIndex) / 6.5 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(8)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Remuneración adicional = Sueldo base / 6.5 * 3 * (Plazas * Factor de plazas) * Periodo
						asAmounts(63)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 0.25 * 48 * asAmounts(4)(iIndex) 'Prima dominical = (Sueldo + turno opcional + percepción adicional) / 30 * 25% * 48 * Plazas
						asAmounts(71)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 8.5 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(17)(1) / 100) + 0.9999999999999) 'Bono de Reyes = SMB * 8.5 * (Plazas * Factor de plazas)
						asAmounts(72)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(18)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(18)(1) / 100) + 0.9999999999999) 'Ayuda compra de útiles = Ayuda compra de útiles * (Plazas * Factor de plazas)
						asAmounts(73)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(20)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(20)(1) / 100) + 0.9999999999999) 'Ayuda de anteojos = Ayuda de anteojos * (Plazas * Factor de plazas)
						asAmounts(74)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(22)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(22)(1) / 100) + 0.9999999999999) 'Ayuda por muerte familiar 1° grado = Ayuda por muerte familiar 1º grado * (Plazas * Factor de plazas)
						asAmounts(75)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(21)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(21)(1) / 100) + 0.9999999999999) 'Impresión de tesis = Ayuda para impresión de tesis * (Plazas * Factor de plazas)
						asAmounts(76)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(23)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(23)(1) / 100) + 0.9999999999999) 'Evento 10 de mayo = Evento 10 de mayo * (Plazas * Factor de plazas)
						asAmounts(77)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(24)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(24)(1) / 100) + 0.9999999999999) 'Evento día del niño = 120 * (Plazas * Factor de plazas)
						asAmounts(78)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 8 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(26)(1) / 100) + 0.9999999999999) 'Evento fomento cultural, turístico y deportivo = SMB * 8 * (Plazas * Factor de plazas)
						asAmounts(79)(iIndex) = asAmounts(5)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(25)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(25)(1) / 100) + 0.9999999999999) 'Evento día del trabajador = Sueldo base * Evento día del trabajador * (Plazas * Factor de Plazas)
						asAmounts(80)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(29)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(29)(1) / 100) + 0.9999999999999) 'Comisión nacional de auxilio = Comisión nacional de auxilio * (Plazas * Factor de plazas)
						asAmounts(82)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(31)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(31)(1) / 100) + 0.9999999999999)  'Premios, estímulos y recompensas = Premio * (Plazas * Factor de plazas)
						'asAmounts(xxx)(iIndex) = 52 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Comedor = 52 * Plazas * Periodo
						asAmounts(84)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(28)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(28)(1) / 100) + 0.9999999999999) 'Becas hijos de trabajadores = Becas hijos de trabajadores * (Plazas * Factor de plazas)
						asAmounts(85)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 12 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) / 2 'Premio 10 de mayo = SMB * 12 * Plazas * Periodo / 2
						asAmounts(86)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 1 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) 'Premio por antigüedad = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * Numero de días a pagar * (Plazas * Factor de plazas)
						asAmounts(87)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 1 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(12)(1) / 100) + 0.9999999999999) 'Días económicos no disfrutados = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * Numero de días a pagar * (Plazas * Factor de plazas)
						asAmounts(88)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(43)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(43)(1) / 100) + 0.9999999999999) 'Premio moneda de oro = Premio * (Plazas * Factor de plazas)
						asAmounts(89)(iIndex) = asAmounts(5)(iIndex) * 0.4 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería = Sueldo * 40% * Plazas * Periodo
						asAmounts(90)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(34)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por asistencia = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas) * Periodo
						asAmounts(91)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(36)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por puntualidad = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * (Plazas * Factor de plazas) * Periodo
						asAmounts(92)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 2 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(36)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Estímulo por desempeño = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * 2 días * (Plazas * Factor de plazas) * Periodo
						asAmounts(93)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex) + asAmounts(39)(iIndex)) / 30 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(37)(1) / 100) + 0.9999999999999) * 4 'Estímulo mérito relevante = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional + prima quinquenal) / 30 * 3 * (Plazas * Factor de plazas) * 4
						asAmounts(94)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(44)(1) / 100) + 0.9999999999999) 'Premio de antigüedad 25 y 30 años = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) * (Plazas * Factor de plazas)
						asAmounts(95)(iIndex) = asAmounts(5)(iIndex) * 0.30 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(45)(1) / 100) + 0.9999999999999) 'Premio trabajador del mes = 30% del sueldo * (Plazas * Factor de plazas)
						asAmounts(99)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(33)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(33)(1) / 100) + 0.9999999999999) 'Aportaciones por servicios de atención para el bienestar y desarrollo infantil = Cuota guardería * plazas * factor sobre plazas
					Case 4 'Enlace
						asAmounts(25)(iIndex) = (asAmounts(5)(iIndex) * 0.2) + (asAmounts(5)(iIndex) * 0.1) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por riesgos profesionales = (Sueldo base * 20%) + (Sueldo base * 10%) * Plazas * Periodo
						asAmounts(29)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(6)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Previsión social múltiple = Previsión social múltiple * Plazas * Periodo
						asAmounts(34)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) * asParameters(CInt(asConditions(1)(iIndex)))(65)(0) / 100 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el seguro de separación individualizado = (Sueldo base + Compensación garantizada) * Factor SSI * Plazas * Periodo
						asAmounts(41)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) / 30 / 8 * 36 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(7)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Horas extras = Sueldo / 30 / 8 * 36 * (Plazas * Factor de plazas) * Periodo
						asAmounts(49)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 3 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Cuotas para el fondo de ahorro del personal civil = SMB * 3 * Plazas * Periodo
						asAmounts(54)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(6)(iIndex)) / 30 * asParameters(CInt(asConditions(1)(iIndex)))(5)(0) * asAmounts(4)(iIndex) 'Ajuste al calendario = (Sueldo + Compensación garantizada) / 30 * Ajuste al calendario * No. Plazas
						asAmounts(57)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(15)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda de transporte = Ayuda de transporte * Plazas * Periodo
						asAmounts(59)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(16)(0) * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Vales de despensa = Vales de despensa * Plazas * Periodo
						asAmounts(61)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 9 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(42)(1) / 100) + 0.9999999999999) 'Premio de aniversario = SMB * 9 * (Plazas * Factor de plazas)
						asAmounts(62)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(40)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por antigüedad = Compensación por antigüedad * (Plazas * Factor de plazas) * Periodo
						asAmounts(69)(iIndex) = asAmounts(5)(iIndex) / 6.5 * 3 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(8)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Remuneración adicional = Sueldo base / 6.5 * 3 * (Plazas * Factor de plazas) * Periodo
						asAmounts(71)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 8.5 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(17)(1) / 100) + 0.9999999999999) 'Bono de Reyes = SMB * 8.5 * (Plazas * Factor de plazas)
						asAmounts(72)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(18)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(18)(1) / 100) + 0.9999999999999) 'Ayuda compra de útiles = Ayuda compra de útiles * (Plazas * Factor de plazas)
						asAmounts(73)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(20)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(20)(1) / 100) + 0.9999999999999) 'Ayuda de anteojos = Ayuda de anteojos * (Plazas * Factor de plazas)
						asAmounts(75)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(21)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(21)(1) / 100) + 0.9999999999999) 'Impresión de tesis = Ayuda para impresión de tesis * (Plazas * Factor de plazas)
						asAmounts(80)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(29)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(29)(1) / 100) + 0.9999999999999) 'Comisión nacional de auxilio = Comisión nacional de auxilio * (Plazas * Factor de plazas)
						asAmounts(82)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(31)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(31)(1) / 100) + 0.9999999999999)  'Premios, estímulos y recompensas = Premio * (Plazas * Factor de plazas)
						asAmounts(85)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(2)(0) * 12 * asAmounts(4)(iIndex) 'Premio 10 de mayo = SMB * 12 * Plazas
						asAmounts(86)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) / 30 * 1 * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(40)(1) / 100) + 0.9999999999999) 'Premio por antigüedad = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) / 30 * Numero de días a pagar * (Plazas * Factor de plazas)
						asAmounts(94)(iIndex) = (asAmounts(5)(iIndex) + asAmounts(25)(iIndex) + asAmounts(62)(iIndex) + asAmounts(68)(iIndex) + asAmounts(69)(iIndex)) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(44)(1) / 100) + 0.9999999999999) 'Premio de antigüedad 25 y 30 años = (Sueldo base + compensación por riesgos profesionales + compensación por antigüedad + turno opcional + percepción adicional) * (Plazas * Factor de plazas)
						asAmounts(99)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(33)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(33)(1) / 100) + 0.9999999999999) 'Aportaciones por servicios de atención para el bienestar y desarrollo infantil = Cuota guardería * plazas * factor sobre plazas
					Case 6 'Becarios
						asAmounts(21)(iIndex) = asAmounts(15)(iIndex) * asAmounts(4)(iIndex) * 3 'Gratificación de becas = Beca a médicos residentes * Plazas * 3
						asAmounts(24)(iIndex) = asAmounts(15)(iIndex) / 30 * 40 * (asParameters(CInt(asConditions(1)(iIndex)))(76)(0) /12) * asAmounts(4)(iIndex) 'Aguinaldo = Beca a médicos residentes / 30 * 40 * (Periodo / 12) * Plazas
						asAmounts(56)(iIndex) = asParameters(CInt(asConditions(1)(iIndex)))(64)(0) * Int((asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(64)(1) / 100) + 0.9999999999999) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda renta a becarios = Ayuda renta becarios * (Plazas * Factor plazas) * Periodo
						asAmounts(60)(iIndex) = asAmounts(15)(iIndex) * 4 * asAmounts(4)(iIndex) 'Material didáctico = Beca a médicos residentes * 4 * Plazas
						asAmounts(96)(iIndex) = asAmounts(24)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(72)(0) / 100 'ISR gratificación de fin de año = Aguinaldo * Factor de ISR (aguinaldo)
				End Select

				asAmounts(36)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al IMSS = xxx * Plazas * Periodo
				asAmounts(37)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones al INFONAVIT = xxx * Plazas * Periodo
				asAmounts(40)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Prima de antigüedad = xxx * Plazas * Periodo
				asAmounts(42)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Diferencias cambiarias = xxx * Plazas * Periodo
				asAmounts(45)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Aportaciones de seguridad social contractuales = xxx * Plazas * Periodo
				asAmounts(50)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Gratificación por renuncia voluntaria = xxx * Plazas * Periodo
				asAmounts(51)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación de bajo desarrollo = xxx * Plazas * Periodo
				asAmounts(52)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Apoyo a la capacitación de los servidores públicos = xxx * Plazas * Periodo
				asAmounts(53)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Compensación por desempeño para la educación indígena = xxx * Plazas * Periodo
				asAmounts(100)(iIndex) = 0 * asAmounts(4)(iIndex) * asParameters(CInt(asConditions(1)(iIndex)))(76)(0) 'Ayuda por servicios = xxx * Plazas * Periodo
				asAmounts(101)(iIndex) = 0 'TOTAL = Suma(12 100)
				For jIndex = 12 To 100
					asAmounts(101)(iIndex) = asAmounts(101)(iIndex) + asAmounts(jIndex)(iIndex)
				Next
			Next

			Select Case lReportID
				Case ISSSTE_1503_REPORTS
					Response.Write "<TABLE BORDER="""
						If Not bForExport Then
							Response.Write "0"
						Else
							Response.Write "1"
						End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
						asColumnsTitles = Split("<SPAN COLS=""7"" />ESTRUCTURA OCUPACIONAL,<SPAN COLS=""7"" />TABULADOR,<SPAN COLS=""8"" />PERCEPCIÓN ORDINARIA,<SPAN COLS=""10"" />PRESTACIONES SOCIALES Y ECONÓMICAS,<SPAN COLS=""9"" />APORTACIONES DE SEGURIDAD SOCIAL Y SEGUROS,<SPAN COLS=""15"" />PARTIDAS DE GASTO CALCULADAS POR DEPENDENCIA,<SPAN COLS=""8"" />OTRAS PARTIDAS ASOCIADAS A LA PLAZA,<SPAN COLS=""34"" />&nbsp;,<SPAN COLS=""5"" />IMPUESTOS DE PARTIDAS,&nbsp;", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If

						asColumnsTitles = Split("Compañía,Zona,Nivel,Código,Puesto,Tipo de tabulador,Plazas,Sueldo base,Comp. garantizada,Comp. por riesgo,Asignación,AGA,Beca,Comp. beca,1103,1201,1202,1204,1326,1509,1512,1325,1103,1204,1204,1305,1306,1322,1326,1507,1508,1511,1401,1401,1403,1404,1407,1408,1410,1411,1413,1301,1302,1319,1329,1406,1409,1412,1414,1414,1415,1501,1505,1512,1513,1702,1103,1306,1507,1507,1511,1512,1512,1702,1301,1305,1308,1319,1319,1319,1319,1319,1409,1507,1507,1507,1507,1507,1507,1507,1507,1512,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1306,1306,1407,1401,1511,&nbsp;", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If

						asColumnsTitles = Split("&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...&nbsp;...Sueldos base...Honorarios...Sueldo base al<BR />personal eventual...Beca a médicos<BR />residentes...Complemento de Beca...Compensación garantizada...Asignación Médica...Aga...Subsidio para el empleo...Gratificación de becas...Retribuciones por Servicios de Carácter Social...Prima vacacional...Aguinaldo...Compensación por Riesgos Profesionales...Compensación a Médicos Residentes...Despensa...Aportaciones al Sistema de Ahorro para el Retiro...Previsión social múltiple...Aportaciones al ISSSTE...Cuota Social...Aportaciones al FOVISSSTE...Seguro institucional...Cuotas para el seguro de separación individualizado...Cuotas para el seguro colectivo de retiro...Aportaciones al IMSS...Aportaciones al INFONAVIT...Aportaciones al Sistema de Ahorro para el Retiro...Prima quinquenal por años de servicios efectivos prestados...Prima de antigüedad...Horas Extras...Diferencias Cambiarias...Cuotas para el seguro de gastos médicos del personal civil...Seguro Responsabilidad Civil...Aportaciones de seguridad social contractuales...Aportaciones al Seguro por Cesantía...Cuota de Cesantía...Depósitos para el ahorro solidario...Cuotas para el fondo de ahorro del personal civil...Gratificación por Renuncia Voluntaria...Compensación de bajo desarrollo... Apoyo a la Capacitación de los Servidores Públicos...Compensación por desempeño para la educación indígena...Ajuste al Calendario...Gratificación de Fin de Año de la Compensación Garantizada...Ayuda Renta a Becarios...Ayuda de Transporte...Apoyo para el Desarrollo y Capacitación...Vales de Despensa...Material Didáctico...Premio de Aniversario...Compensación por Antigüedad...Prima Dominical...Remuneraciones por Suplencias...Remuneraciones por Guardias...Servicios Prioritarios de Atención Primaria a la Salud...Rezago Quirurgico...Turno Opcional...Remuneración Adicional...Cuota para el Seguro de Responsabilidad Civil y de Responsabilidad Profesional para Médicos y Enfermeras...Bono de Reyes...Ayuda Compra de Útiles...Ayuda de Anteojos...Ayuda por muerte familiar 1° grado...Impresión de Tesis...Evento 10 de Mayo...Evento día del niño...Evento Fomento Cultural, Turístico y Deportivo...Evento día del Trabajador...Comisión Nacional de Auxilio...Jornada nocturna adicional por día festivo...Premios, Estímulos y Recompensas...Participación de Inventarios Físicos...Becas hijos de trabajadores...Premio 10 de Mayo...Premio por  Antigüedad...Días económicos no disfrutados...Premio Moneda de Oro...Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería...Estímulo por asistencia...Estímulo de Puntualidad...Estímulo de Desempeño...Estímulo Mérito Relevante...Premio de Antigüedad 25 y 30 años...Premio Trabajador del mes...ISR Gratificación de Fin de AÑO...ISR Gratificación de Fin de Año de la Compensación Garantizada...ISR Seguro de Separación Individualizado...Aportaciones por servicios de atencion para el bienestar y desarrollo infantil...Ayuda por Servicios...TOTAL", "...", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If

						asCellAlignments = Split(",CENTER,CENTER,,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
						For jIndex = 0 To UBound(asAmounts(0))
							sRowContents = ""
							sRowContents = asAmounts(0)(jIndex)
							For iIndex = 5 To UBound(asAmounts)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asAmounts(iIndex)(jIndex), 2, True, False, True)
							Next
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Next

						sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
						dTotal = 0
						For jIndex = 0 To UBound(asAmounts(0))
							dTotal = dTotal + asAmounts(4)(jIndex)
						Next
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 0, True, False, True) & "</B>"
						For iIndex = 5 To 11
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						Next
						For iIndex = 12 To UBound(asAmounts)
							dTotal = 0
							For jIndex = 0 To UBound(asAmounts(0))
								dTotal = dTotal + asAmounts(iIndex)(jIndex)
							Next
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
						Next
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					Response.Write "</TABLE>" & vbNewLine
				Case ISSSTE_1561_REPORTS
				Case ISSSTE_1562_REPORTS
				Case ISSSTE_1563_REPORTS
					If bForExport Then Response.Write Replace(GetFileContents(Server.MapPath("Templates\HeaderForReport_1563.htm"), sErrorDescription), "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					Response.Write "<TABLE BORDER="""
						If Not bForExport Then
							Response.Write "0"
						Else
							Response.Write "1"
						End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
						asColumnsTitles = Split("Compañía,Zona,Nivel,Código,Puesto,Tipo de tabulador,Plazas,Sueldo base,Sueldo base<BR />Colectivo por periodo,Comp. garantizada,Comp. garantizada<BR />Colectiva por periodo,Comp. por riesgo,Comp. por riesgo<BR />Colectiva por periodo,Asignación,Asignación<BR />Colectiva por periodo,AGA,AGA<BR />Colectiva por periodo,Beca,Beca<BR />Colectiva por periodo,Comp. beca,Comp. beca<BR />Colectiva por periodo", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If

						asRowsTitles = Split(",,,,,5,12,6,17,7,25,8,18,9,19,10,15,11,16", ",")
						asCellAlignments = Split(",CENTER,CENTER,,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
						For jIndex = 0 To UBound(asAmounts(0))
							sRowContents = ""
							sRowContents = asAmounts(0)(jIndex)
							For iIndex = 5 To UBound(asRowsTitles)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asAmounts(CInt(asRowsTitles(iIndex)))(jIndex), 2, True, False, True)
							Next
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Next

						sRowContents = "<SPAN COLS=""6"" /><B>TOTAL</B>"
						dTotal = 0
						For jIndex = 0 To UBound(asAmounts(0))
							dTotal = dTotal + asAmounts(4)(jIndex)
						Next
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 0, True, False, True) & "</B>"
						For iIndex = 5 To UBound(asRowsTitles)
							dTotal = 0
							For jIndex = 0 To UBound(asAmounts(0))
								dTotal = dTotal + asAmounts(CInt(asRowsTitles(iIndex)))(jIndex)
							Next
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
						Next
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					Response.Write "</TABLE><BR />" & vbNewLine

					Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
						Response.Write "<TD COLSPAN=""7"">&nbsp;</TD>"
						Response.Write "<TD><TABLE BORDER="""
							If Not bForExport Then
								Response.Write "0"
							Else
								Response.Write "1"
							End If
						Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
							asRowsTitles = Split(",", ",")
							asRowsTitles(0) = Split("....................................Sueldos base...Honorarios...Sueldo base al<BR />personal eventual...Beca a médicos<BR />residentes...Complemento de Beca...Compensación garantizada...Asignación Médica...Aga...Subsidio para el empleo...Gratificación de becas...Retribuciones por Servicios de Carácter Social...Prima vacacional...Aguinaldo...Compensación por Riesgos Profesionales...Compensación a Médicos Residentes...Despensa...Aportaciones al Sistema de Ahorro para el Retiro...Previsión social múltiple...Aportaciones al ISSSTE...Cuota Social...Aportaciones al FOVISSSTE...Seguro institucional...Cuotas para el seguro de separación individualizado...Cuotas para el seguro colectivo de retiro...Aportaciones al IMSS...Aportaciones al INFONAVIT...Aportaciones al Sistema de Ahorro para el Retiro...Prima quinquenal por años de servicios efectivos prestados...Prima de antigüedad...Horas Extras...Diferencias Cambiarias...Cuotas para el seguro de gastos médicos del personal civil...Seguro Responsabilidad Civil...Aportaciones de seguridad social contractuales...Aportaciones al Seguro por Cesantía...Cuota de Cesantía...Depósitos para el ahorro solidario...Cuotas para el fondo de ahorro del personal civil...Gratificación por Renuncia Voluntaria...Compensación de bajo desarrollo... Apoyo a la Capacitación de los Servidores Públicos...Compensación por desempeño para la educación indígena...Ajuste al Calendario...Gratificación de Fin de Año de la Compensación Garantizada...Ayuda Renta a Becarios...Ayuda de Transporte...Apoyo para el Desarrollo y Capacitación...Vales de Despensa...Material Didáctico...Premio de Aniversario...Compensación por Antigüedad...Prima Dominical...Remuneraciones por Suplencias...Remuneraciones por Guardias...Servicios Prioritarios de Atención Primaria a la Salud...Rezago Quirurgico...Turno Opcional...Remuneración Adicional...Cuota para el Seguro de Responsabilidad Civil y de Responsabilidad Profesional para Médicos y Enfermeras...Bono de Reyes...Ayuda Compra de Útiles...Ayuda de Anteojos...Ayuda por muerte familiar 1° grado...Impresión de Tesis...Evento 10 de Mayo...Evento día del niño...Evento Fomento Cultural, Turístico y Deportivo...Evento día del Trabajador...Comisión Nacional de Auxilio...Jornada nocturna adicional por día festivo...Premios, Estímulos y Recompensas...Participación de Inventarios Físicos...Becas hijos de trabajadores...Premio 10 de Mayo...Premio por  Antigüedad...Días económicos no disfrutados...Premio Moneda de Oro...Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería...Estímulo por asistencia...Estímulo de Puntualidad...Estímulo de Desempeño...Estímulo Mérito Relevante...Premio de Antigüedad 25 y 30 años...Premio Trabajador del mes...ISR Gratificación de Fin de AÑO...ISR Gratificación de Fin de Año de la Compensación Garantizada...ISR Seguro de Separación Individualizado...Aportaciones por servicios de atencion para el bienestar y desarrollo infantil...Ayuda por Servicios...TOTAL", "...", -1, vbBinaryCompare)
							asRowsTitles(1) = Split(",,,,,,,,,,,,1103,1201,1202,1204,1326,1509,1512,1325,1103,1204,1204,1305,1306,1322,1326,1507,1508,1511,1401,1401,1403,1404,1407,1408,1410,1411,1413,1301,1302,1319,1329,1406,1409,1412,1414,1414,1415,1501,1505,1512,1513,1702,1103,1306,1507,1507,1511,1512,1512,1702,1301,1305,1308,1319,1319,1319,1319,1319,1409,1507,1507,1507,1507,1507,1507,1507,1507,1512,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1702,1306,1306,1407,1401,1511,", ",", -1, vbBinaryCompare)
							asColumnsTitles = Split("Capítulo<BR />Concepto,Periodo<BR />Colectivo,Complemento<BR />Colectivo,Total<BR />&nbsp;", ",", -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
							Else
								If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
									lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								Else
									lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
								End If
							End If

							asCellAlignments = Split(",RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
							For iIndex = 12 To UBound(asAmounts) - 1
								dTotal = 0
								For jIndex = 0 To UBound(asAmounts(0))
									dTotal = dTotal + asAmounts(iIndex)(jIndex)
								Next
								sRowContents = ""
								If dTotal > 0 Then
									sRowContents = asRowsTitles(0)(iIndex) & "&nbsp;" & asRowsTitles(1)(iIndex) & TABLE_SEPARATOR & FormatNumber(dTotal, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber((dTotal / lPeriods * (12 - lPeriods)), 2, True, False, True) & TABLE_SEPARATOR & FormatNumber((dTotal / lPeriods * 12), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
								End If
							Next
							dTotal = 0
							For jIndex = 0 To UBound(asAmounts(0))
								dTotal = dTotal + asAmounts(iIndex)(jIndex)
							Next
							sRowContents = "<B>" & asRowsTitles(0)(iIndex) & "&nbsp;" & asRowsTitles(1)(iIndex) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>0.00</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Response.Write "</TABLE></TD>" & vbNewLine
					Response.Write "</TR></TABLE>" & vbNewLine
			End Select
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1503 = lErrorNumber
	Err.Clear
End Function

Function DisplayReport1503Parameters(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the parameters for budget simulation
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReport1503Parameters"
	Dim oItem
	Dim asItem
	Dim asParameters
	Dim iIndex
	Dim jIndex
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments

	asParameters = Split(",,,,,,", ",")
	'Médica, paramédica y grupos afines
	asParameters(0) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "150", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "5", "", ""),_
		Array("Previsión social múltiple", "150", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "20", ""),_
		Array("Remuneración adicional", "", "20", ""),_
		Array("Remuneración por guardias", "", "20", ""),_
		Array("Remuneración por suplencias", "", "20", ""),_
		Array("Turno opcional", "", "20", ""),_
		Array("Días económicos no disfrutados", "", "20", ""),_
		Array("Prima dominical", "", "20", ""),_
		Array("Jornada nocturna por día festivo", "", "20", ""),_
		Array("Ayuda de transporte", "120", "", ""),_
		Array("Vales de despensa", "8450", "", ""),_
		Array("Bono de Reyes", "", "100", ""),_
		Array("Ayuda para compra de útiles", "110", "100", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "1400", "40", ""),_
		Array("Ayuda para impresión de tesis", "3500", "10", ""),_
		Array("Ayuda por muerte familiar 1º grado", "2000", "5", ""),_
		Array("Evento 10 de mayo", "120", "50", ""),_
		Array("Evento del día del niño", "120", "60", ""),_
		Array("Evento del día del trabajador", "5", "80", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "20", ""),_
		Array("Participación en inventarios físicos", "", "20", ""),_
		Array("Becas para hijos de trabajadores", "800", "15", ""),_
		Array("Comisión nacional de auxilio", "1333.33", "20", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "8500", "20", ""),_
		Array("Premio 10 de mayo", "", "50", ""),_
		Array("Cuota guardería", "29700", "50", ""),_
		Array("Estímulo por asistencia", "", "20", ""),_
		Array("Estímulo de puntualidad", "", "20", ""),_
		Array("Estímulo de desempeño", "", "20", ""),_
		Array("Estímulo de mérito relevante", "", "20", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "20", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "471", "10", ""),_
		Array("Factor prima antigüedad", "0", "10", "%"),_
		Array("Premio de aniversario", "", "100", ""),_
		Array("Premio moneda de oro", "6892.53", "10", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "10", ""),_
		Array("Premio del trabajador del mes", "", "10", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "59.13", "", ""),_
		Array("Factor SAR", "0", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "73", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "20", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "54.50", "", ""),_
		Array("Rezago quirúrgico", "1000", "100", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "0", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "0", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "15", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "0", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Mandos = Funcionarios
	asParameters(1) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "77", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "0", "", ""),_
		Array("Previsión social múltiple", "150", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "0", ""),_
		Array("Remuneración adicional", "", "0", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "0", ""),_
		Array("Días económicos no disfrutados", "", "0", ""),_
		Array("Prima dominical", "", "0", ""),_
		Array("Jornada nocturna por día festivo", "", "0", ""),_
		Array("Ayuda de transporte", "120", "", ""),_
		Array("Vales de despensa", "0", "", ""),_
		Array("Bono de Reyes", "", "0", ""),_
		Array("Ayuda para compra de útiles", "0", "0", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "0", "0", ""),_
		Array("Ayuda para impresión de tesis", "0", "0", ""),_
		Array("Ayuda por muerte familiar 1º grado", "0", "0", ""),_
		Array("Evento 10 de mayo", "0", "0", ""),_
		Array("Evento del día del niño", "0", "0", ""),_
		Array("Evento del día del trabajador", "0", "0", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "0", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "0", "0", ""),_
		Array("Comisión nacional de auxilio", "0", "0", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "0", "0", ""),_
		Array("Premio 10 de mayo", "", "0", ""),_
		Array("Cuota guardería", "0", "0", ""),_
		Array("Estímulo por asistencia", "", "0", ""),_
		Array("Estímulo de puntualidad", "", "0", ""),_
		Array("Estímulo de desempeño", "", "0", ""),_
		Array("Estímulo de mérito relevante", "", "0", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "471", "0", ""),_
		Array("Factor prima antigüedad", "0", "0", "%"),_
		Array("Premio de aniversario", "", "0", ""),_
		Array("Premio moneda de oro", "0", "0", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "0", ""),_
		Array("Premio del trabajador del mes", "", "0", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "2", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "110", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "0", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "", "", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "8.6", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "30", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "30", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "30", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Operativos
	asParameters(2) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "150", "", ""),_
		Array("CODECA", "800", "", ""),_
		Array("Ajuste al calendario", "5", "", ""),_
		Array("Previsión social múltiple", "150", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "20", ""),_
		Array("Remuneración adicional", "", "20", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "20", ""),_
		Array("Días económicos no disfrutados", "", "20", ""),_
		Array("Prima dominical", "", "20", ""),_
		Array("Jornada nocturna por día festivo", "", "20", ""),_
		Array("Ayuda de transporte", "120", "", ""),_
		Array("Vales de despensa", "8450", "", ""),_
		Array("Bono de Reyes", "", "100", ""),_
		Array("Ayuda para compra de útiles", "110", "100", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "1400", "40", ""),_
		Array("Ayuda para impresión de tesis", "3500", "10", ""),_
		Array("Ayuda por muerte familiar 1º grado", "2000", "5", ""),_
		Array("Evento 10 de mayo", "120", "50", ""),_
		Array("Evento del día del niño", "120", "60", ""),_
		Array("Evento del día del trabajador", "5", "80", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "20", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "800", "15", ""),_
		Array("Comisión nacional de auxilio", "1333.33", "20", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "8500", "20", ""),_
		Array("Premio 10 de mayo", "", "50", ""),_
		Array("Cuota guardería", "29700", "50", ""),_
		Array("Estímulo por asistencia", "", "20", ""),_
		Array("Estímulo de puntualidad", "", "20", ""),_
		Array("Estímulo de desempeño", "", "20", ""),_
		Array("Estímulo de mérito relevante", "", "20", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "471", "10", ""),_
		Array("Factor prima antigüedad", "0", "0", "%"),_
		Array("Premio de aniversario", "", "100", ""),_
		Array("Premio moneda de oro", "6892.53", "10", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "10", ""),_
		Array("Premio del trabajador del mes", "", "10", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "2", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "73", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "20", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "", "", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "0", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "0", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "15", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "15", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Alta Responsabilidad
	asParameters(3) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "77", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "0", "", ""),_
		Array("Previsión social múltiple", "150", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "0", ""),_
		Array("Remuneración adicional", "", "0", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "0", ""),_
		Array("Días económicos no disfrutados", "", "0", ""),_
		Array("Prima dominical", "", "0", ""),_
		Array("Jornada nocturna por día festivo", "", "0", ""),_
		Array("Ayuda de transporte", "120", "", ""),_
		Array("Vales de despensa", "0", "", ""),_
		Array("Bono de Reyes", "", "0", ""),_
		Array("Ayuda para compra de útiles", "0", "0", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "0", "0", ""),_
		Array("Ayuda para impresión de tesis", "0", "0", ""),_
		Array("Ayuda por muerte familiar 1º grado", "0", "0", ""),_
		Array("Evento 10 de mayo", "0", "0", ""),_
		Array("Evento del día del niño", "0", "0", ""),_
		Array("Evento del día del trabajador", "0", "0", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "0", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "0", "0", ""),_
		Array("Comisión nacional de auxilio", "0", "0", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "0", "0", ""),_
		Array("Premio 10 de mayo", "", "0", ""),_
		Array("Cuota guardería", "0", "0", ""),_
		Array("Estímulo por asistencia", "", "0", ""),_
		Array("Estímulo de puntualidad", "", "0", ""),_
		Array("Estímulo de desempeño", "", "0", ""),_
		Array("Estímulo de mérito relevante", "", "0", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "471", "0", ""),_
		Array("Factor prima antigüedad", "0", "0", "%"),_
		Array("Premio de aniversario", "", "0", ""),_
		Array("Premio moneda de oro", "0", "0", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "0", ""),_
		Array("Premio del trabajador del mes", "", "0", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "2", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "110", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "0", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "", "", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "8.6", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "30", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "30", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "30", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Enlaces
	asParameters(4) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "77", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "5", "", ""),_
		Array("Previsión social múltiple", "95.58", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "20", ""),_
		Array("Remuneración adicional", "", "20", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "0", ""),_
		Array("Días económicos no disfrutados", "", "0", ""),_
		Array("Prima dominical", "", "0", ""),_
		Array("Jornada nocturna por día festivo", "", "0", ""),_
		Array("Ayuda de transporte", "80.24", "", ""),_
		Array("Vales de despensa", "8450", "", ""),_
		Array("Bono de Reyes", "", "100", ""),_
		Array("Ayuda para compra de útiles", "110", "100", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "1400", "40", ""),_
		Array("Ayuda para impresión de tesis", "3500", "10", ""),_
		Array("Ayuda por muerte familiar 1º grado", "2000", "5", ""),_
		Array("Evento 10 de mayo", "0", "0", ""),_
		Array("Evento del día del niño", "0", "0", ""),_
		Array("Evento del día del trabajador", "0", "0", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "0", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "0", "0", ""),_
		Array("Comisión nacional de auxilio", "0", "0", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "0", "0", ""),_
		Array("Premio 10 de mayo", "", "50", ""),_
		Array("Cuota guardería", "0", "0", ""),_
		Array("Estímulo por asistencia", "", "0", ""),_
		Array("Estímulo de puntualidad", "", "0", ""),_
		Array("Estímulo de desempeño", "", "0", ""),_
		Array("Estímulo de mérito relevante", "", "0", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "471", "10", ""),_
		Array("Factor prima antigüedad", "0", "10", "%"),_
		Array("Premio de aniversario", "", "100", ""),_
		Array("Premio moneda de oro", "6892.53", "10", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "10", ""),_
		Array("Premio del trabajador del mes", "", "0", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "2", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "73", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "0", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "", "", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "8.6", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "30", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "20", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "20", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Residentes
	asParameters(5) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "150", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "0", "", ""),_
		Array("Previsión social múltiple", "0", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "0", ""),_
		Array("Remuneración adicional", "", "0", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "0", ""),_
		Array("Días económicos no disfrutados", "", "0", ""),_
		Array("Prima dominical", "", "0", ""),_
		Array("Jornada nocturna por día festivo", "", "0", ""),_
		Array("Ayuda de transporte", "120", "", ""),_
		Array("Vales de despensa", "8450", "", ""),_
		Array("Bono de Reyes", "", "100", ""),_
		Array("Ayuda para compra de útiles", "110", "100", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "1400", "40", ""),_
		Array("Ayuda para impresión de tesis", "3500", "10", ""),_
		Array("Ayuda por muerte familiar 1º grado", "2000", "5", ""),_
		Array("Evento 10 de mayo", "0", "0", ""),_
		Array("Evento del día del niño", "0", "0", ""),_
		Array("Evento del día del trabajador", "0", "0", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "0", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "800", "15", ""),_
		Array("Comisión nacional de auxilio", "1333.33", "0", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "0", "0", ""),_
		Array("Premio 10 de mayo", "", "50", ""),_
		Array("Cuota guardería", "0", "0", ""),_
		Array("Estímulo por asistencia", "", "0", ""),_
		Array("Estímulo de puntualidad", "", "0", ""),_
		Array("Estímulo de desempeño", "", "0", ""),_
		Array("Estímulo de mérito relevante", "", "0", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "109", "10", ""),_
		Array("Compensación por antigüedad", "0", "0", ""),_
		Array("Factor prima antigüedad", "0", "0", "%"),_
		Array("Premio de aniversario", "", "100", ""),_
		Array("Premio moneda de oro", "0", "0", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "0", ""),_
		Array("Premio del trabajador del mes", "", "0", ""),_
		Array("Subsidio para el empleo", "13", "", "%"),_
		Array("Cuota social", "259.37", "", ""),_
		Array("Cuota ISSSTE", "3,112.44", "", ""),_
		Array("Factor ISSSTE", "9.97", "", "%"),_
		Array("Factor FOVISSSTE", "5", "", "%"),_
		Array("Factor ahorro solidario", "6.5", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "2", "", "%"),_
		Array("Cuota cesantía", "97.77", "", ""),_
		Array("Factor cesantía", "0.3175", "", ""),_
		Array("Seguro colectivo", "59.13", "", ""),_
		Array("Seguro de responsabilidad civil", "73", "", ""),_
		Array("Seguro de vida del personal civil", "2.29", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "0", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "0", "0", ""),_
		Array("Ayuda renta becarios", "0", "0", ""),_
		Array("Factor SSI", "0", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "0", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "15", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "0", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)
	'Becarios
	asParameters(6) = Array(_
		Array("SMG", "57.46", "", ""),_
		Array("Tope SMG", "119.53", "", ""),_
		Array("SMB", "119.53", "", ""),_
		Array("Despensa", "11.70", "", ""),_
		Array("CODECA", "0", "", ""),_
		Array("Ajuste al calendario", "0", "", ""),_
		Array("Previsión social múltiple", "0", "", ""),_
		Array("Remuneraciones por horas extraordinarias", "", "0", ""),_
		Array("Remuneración adicional", "", "0", ""),_
		Array("Remuneración por guardias", "", "0", ""),_
		Array("Remuneración por suplencias", "", "0", ""),_
		Array("Turno opcional", "", "0", ""),_
		Array("Días económicos no disfrutados", "", "0", ""),_
		Array("Prima dominical", "", "0", ""),_
		Array("Jornada nocturna por día festivo", "", "0", ""),_
		Array("Ayuda de transporte", "0", "", ""),_
		Array("Vales de despensa", "8450", "", ""),_
		Array("Bono de Reyes", "", "0", ""),_
		Array("Ayuda para compra de útiles", "0", "0", ""),_
		Array("Material didáctico", "", "", ""),_
		Array("Ayuda para anteojos", "0", "0", ""),_
		Array("Ayuda para impresión de tesis", "0", "0", ""),_
		Array("Ayuda por muerte familiar 1º grado", "0", "0", ""),_
		Array("Evento 10 de mayo", "0", "0", ""),_
		Array("Evento del día del niño", "0", "0", ""),_
		Array("Evento del día del trabajador", "0", "0", "%"),_
		Array("Evento fomento cultural, turístico y deportivo", "", "0", ""),_
		Array("Participación en inventarios físicos", "", "0", ""),_
		Array("Becas para hijos de trabajadores", "0", "0", ""),_
		Array("Comisión nacional de auxilio", "0", "0", ""),_
		Array("Capacitación de los servidores públicos", "", "", ""),_
		Array("Premios, estímulos y recompensas", "0", "0", ""),_
		Array("Premio 10 de mayo", "", "0", ""),_
		Array("Cuota guardería", "0", "0", ""),_
		Array("Estímulo por asistencia", "", "0", ""),_
		Array("Estímulo de puntualidad", "", "0", ""),_
		Array("Estímulo de desempeño", "", "0", ""),_
		Array("Estímulo de mérito relevante", "", "0", ""),_
		Array("Estímulos a la productividad, eficiencia y calidad a favor del personal médico y de enfermería", "", "0", ""),_
		Array("Cuota de prima quinquenal", "0", "0", ""),_
		Array("Compensación por antigüedad", "0", "0", ""),_
		Array("Factor prima antigüedad", "0", "0", "%"),_
		Array("Premio de aniversario", "", "0", ""),_
		Array("Premio moneda de oro", "0", "0", ""),_
		Array("Premio de antigüedad 25 y 30 años", "", "0", ""),_
		Array("Premio del trabajador del mes", "", "0", ""),_
		Array("Subsidio para el empleo", "0", "", "%"),_
		Array("Cuota social", "0", "", ""),_
		Array("Cuota ISSSTE", "0", "", ""),_
		Array("Factor ISSSTE", "0", "", "%"),_
		Array("Factor FOVISSSTE", "0", "", "%"),_
		Array("Factor ahorro solidario", "0", "", "%"),_
		Array("Factor fondo de pensiones", "", "", "%"),_
		Array("Cuota SAR", "0", "", ""),_
		Array("Factor SAR", "0", "", "%"),_
		Array("Cuota cesantía", "0", "", ""),_
		Array("Factor cesantía", "0", "", ""),_
		Array("Seguro colectivo", "23.80", "", ""),_
		Array("Seguro de responsabilidad civil", "0", "", ""),_
		Array("Seguro de vida del personal civil", "0", "", "%"),_
		Array("Retribuciones por servicios de carácter social", "0", "0", ""),_
		Array("Servicios prioritarios de atención a la salud", "", "0", ""),_
		Array("Cuota para el seguro de responsabilidad civil y de responsabilidad profesional para médicos y enfermeras", "0", "", ""),_
		Array("Rezago quirúrgico", "0", "0", ""),_
		Array("Ayuda renta becarios", "89.26", "100", ""),_
		Array("Factor SSI", "0", "", "%"),_
		Array("Factor NSI", "0", "", "%"),_
		Array("Cuota de SGMM (titular)", "0", "", ""),_
		Array("Cuota de SGMM (cónyuge o concubina)", "0", "", ""),_
		Array("Cuota de SGMM (hijos)", "0", "", ""),_
		Array("Número de beneficiarios SGMM (número de hijos)", "0", "", ""),_
		Array("Factor de ISR (SSI)", "0", "", "%"),_
		Array("Factor de ISR (aguinaldo, prima vacacional)", "15", "", "%"),_
		Array("Factor de ISR (aguinaldo de la compensación)", "0", "", "%"),_
		Array("Tipo de cambio", "0", "", ""),_
		Array("<B>Previsión incremento " & Year(Date()) & "</B>", "0", "", "%"),_
		Array("<B>BASE DE CÁLCULO. PERIODO</B>", "12", "", "")_
	)

	For Each oItem In oRequest
		If InStr(1, oItem, "P_", vbBinaryCompare) = 1 Then
			asItem = Split(oItem, "_")
			asParameters(CInt(asItem(1)))(CInt(asItem(2)))(CInt(asItem(3)) + 1) = oRequest(oItem).Item
		End If
	Next
	If bForExport Then
	Else
		Call GetConditionFromURL(oRequest, sCondition, -1, -1)
		sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Positions.", "BudgetsPositions."), "Companies.", "BudgetsPositions."), "Employees.", "BudgetsPositions."), "GroupGradeLevels.", "BudgetsPositions."), "BudgetsBudgetsPositions.", "BudgetsPositions.")
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeeTypes.EmployeeTypeID, EmployeeTypes.EmployeeTypeName From BudgetsPositions, EmployeeTypes Where (BudgetsPositions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeeTypes.EndDate=30000000) And (EmployeeTypes.EmployeeTypeID>-1) " & sCondition & " Order By EmployeeTypes.EmployeeTypeName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
				Do While Not oRecordset.EOF
					Response.Write "<TD VALIGN=""TOP""><TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
						asCellWidths = Split("200,100,10", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,RIGHT", ",", -1, vbBinaryCompare)
						asColumnsTitles = Split("<SPAN COLS=""3"" />" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)), ";;;", -1, vbBinaryCompare)
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
						asColumnsTitles = Split("Parámetro,Monto,Factor<BR />plazas", ",", -1, vbBinaryCompare)
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If

						For iIndex = 0 To UBound(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value)))
							If (Len(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(0)) > 0) And ((Len(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(1)) > 0) Or (Len(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(2)) > 0)) Then
								sRowContents = asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(0) & TABLE_SEPARATOR
								If Len(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(1)) > 0 Then
									sRowContents = sRowContents & "<INPUT TYPE=""TEXT"" NAME=""P_" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "_" & iIndex & "_0"" ID=""P_" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "_" & iIndex & "_0Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(1) & """ CLASS=""TextFields"" />" & asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(3)
								Else
									sRowContents = sRowContents & "&nbsp;"
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR
								If Len(asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(2)) > 0 Then
									sRowContents = sRowContents & "<INPUT TYPE=""TEXT"" NAME=""P_" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "_" & iIndex & "_1"" ID=""P_" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "_" & iIndex & "_0Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & asParameters(CInt(oRecordset.Fields("EmployeeTypeID").Value))(iIndex)(2) & """ CLASS=""TextFields"" />%"
								Else
									sRowContents = sRowContents & "&nbsp;"
								End If
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						Next
					Response.Write "</TABLE></TD>" & vbNewLine
					Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>" & vbNewLine
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TR></TABLE>" & vbNewLine
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayReport1503Parameters = lErrorNumber
	Err.Clear
End Function

Function DisplayReport1503Positions(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the positions for budget simulation
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReport1503Positions"
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(Replace(sCondition, "Employees", "BudgetsPositions"), "EconomicZones", "BudgetsPositions"), "Companies", "BudgetsPositions")
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetsPositions.PositionID, PositionShortName, PositionName, BudgetsPositions.EmployeeTypeID, EconomicZoneID, TotalPositions, CompanyName, LevelName, GroupGradeLevelName, EmployeeTypeShortName, EmployeeTypeName From BudgetsPositions, Companies, Levels, GroupGradeLevels, EmployeeTypes Where (BudgetsPositions.CompanyID=Companies.CompanyID) And (BudgetsPositions.LevelID=Levels.LevelID) And (BudgetsPositions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (BudgetsPositions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (BudgetsPositions.EndDate=30000000) " & sCondition & " Order By PositionShortName, EmployeeTypeShortName, CompanyName, EconomicZoneID, LevelName, GroupGradeLevelName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Compañía,Zona,Nivel,Código,Puesto,Tipo de tabulador,Plazas", ",", -1, vbBinaryCompare)
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
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
					If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value) & """ ID=""TotalPositions_" & CStr(oRecordset.Fields("PositionID").Value) & "Txt"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & CStr(oRecordset.Fields("TotalPositions").Value) & """ CLASS=""TextFields"" />"
					sRowContents = sRowContents & "<POSITION_ID_" & CStr(oRecordset.Fields("PositionID").Value) & " />"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If


	oRecordset.Close
	Set oRecordset = Nothing
	DisplayReport1503Positions = lErrorNumber
	Err.Clear
End Function

Function DisplayReport1503Saved(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the saved reports for the budget
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReport1503Saved"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener el listado de los costeos guardados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Reports.*, UserName, UserLastName From Reports, Users Where (Reports.UserID=Users.UserID) And (Reports.ConstantID=1503) And (ReportDescription<>'" & CATALOG_SEPARATOR & "') Order By ReportName, ReportID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("&nbsp;,Nombre,Descripción,Creado por", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = "<INPUT TYPE=""RADIO"" NAME=""RecordID"" ID=""RecordIDRd"" VALUE=""" & CStr(oRecordset.Fields("ReportID").Value) & """ />"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReportName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("ReportDescription").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value))

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen costeos guardados en el sistema."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayReport1503Saved = lErrorNumber
	Err.Clear
End Function

Function BuildReport1504(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the money records for the budget
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1504"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim iIndex
	Dim asTotals
	Dim bFirst
	Dim sTotals
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim sTempRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	If Len(Trim(sTableNames)) > 0 Then sTableNames = ", " & sTableNames
	sErrorDescription = "No se pudo obtener el presupuesto original y el presupuesto modificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(OriginalAmount) As TotalOriginalAmount,  Sum(ModifiedAmount) As TotalModifiedAmount, BudgetMonth " & sFieldNames & " From BudgetsMoney " & sTableNames & " Where (BudgetsMoney.BudgetYear>0) " & sCondition & sJoinCondition & " Group By " & sSortFields & ", BudgetMonth Order By " & sSortFields & ", BudgetMonth", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, 0, "Mes,Original,Modificado,Ejercido,Restante"), ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asTotals = "0,0,0,0,0,0,0,0,0,0,0,0,0;0,0,0,0,0,0,0,0,0,0,0,0,0;0,0,0,0,0,0,0,0,0,0,0,0,0;0,0,0,0,0,0,0,0,0,0,0,0,0"
				asTotals = Split(asTotals, ";")
				asTotals(0) = Split(asTotals(0), ",")
				asTotals(1) = Split(asTotals(1), ",")
				asTotals(2) = Split(asTotals(2), ",")
				asTotals(3) = Split(asTotals(3), ",")
				For iIndex = 0 to 12
					asTotals(0)(iIndex) = 0
					asTotals(1)(iIndex) = 0
					asTotals(2)(iIndex) = 0
					asTotals(3)(iIndex) = 0
				Next
				bFirst = True
				sRowContents = ""
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 3, -1, sTempRowContents)
					If StrComp(sTempRowContents, sRowContents, vbBinaryCompare) <> 0 Then
						If Not bFirst Then
							asTotals(0)(0) = 0
							asTotals(1)(0) = 0
							For iIndex = 1 To 12
								If (Len(oRequest("BudgetMonth").Item) = 0) Or (InStr(1, ("," & Replace(oRequest("BudgetMonth").Item, " ", "") & ","), ("," & iIndex & ","), vbBinaryCompare) > 0) Then
									sTotals = TABLE_SEPARATOR & asMonthNames_es(iIndex) & TABLE_SEPARATOR & FormatNumber(asTotals(0)(iIndex), 2, True, False, True)
									asTotals(0)(0) = asTotals(0)(0) + asTotals(0)(iIndex)
									asTotals(2)(iIndex) = asTotals(2)(iIndex) + asTotals(0)(iIndex)
									asTotals(0)(iIndex) = 0
									sTotals = sTotals & TABLE_SEPARATOR & FormatNumber(asTotals(1)(iIndex), 2, True, False, True)
									asTotals(1)(0) = asTotals(1)(0) + asTotals(1)(iIndex)
									asTotals(3)(iIndex) = asTotals(3)(iIndex) + asTotals(1)(iIndex)
									asTotals(1)(iIndex) = 0
									asRowContents = Split(sRowContents & sTotals, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
								End If
							Next
							sRowContents = sRowContents & TABLE_SEPARATOR & "Total" & TABLE_SEPARATOR & FormatNumber(asTotals(0)(0), 2, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(1)(0), 2, True, False, True)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
						bFirst = False
						lErrorNumber = BuildRowData(oRecordset, 3, -1, sRowContents)
					End If
					asTotals(0)(CInt(oRecordset.Fields("BudgetMonth").Value)) = CDbl(oRecordset.Fields("TotalOriginalAmount").Value)
					asTotals(1)(CInt(oRecordset.Fields("BudgetMonth").Value)) = CDbl(oRecordset.Fields("TotalModifiedAmount").Value)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				asTotals(0)(0) = 0
				asTotals(1)(0) = 0
				For iIndex = 1 To 12
					If (Len(oRequest("BudgetMonth").Item) = 0) Or (InStr(1, ("," & Replace(oRequest("BudgetMonth").Item, " ", "") & ","), ("," & iIndex & ","), vbBinaryCompare) > 0) Then
						sTotals = TABLE_SEPARATOR & asMonthNames_es(iIndex) & TABLE_SEPARATOR & FormatNumber(asTotals(0)(iIndex), 2, True, False, True)
						asTotals(0)(0) = asTotals(0)(0) + asTotals(0)(iIndex)
						asTotals(2)(iIndex) = asTotals(2)(iIndex) + asTotals(0)(iIndex)
						asTotals(0)(iIndex) = 0
						sTotals = sTotals & TABLE_SEPARATOR & FormatNumber(asTotals(1)(iIndex), 2, True, False, True)
						asTotals(1)(0) = asTotals(1)(0) + asTotals(1)(iIndex)
						asTotals(3)(iIndex) = asTotals(3)(iIndex) + asTotals(1)(iIndex)
						asTotals(1)(iIndex) = 0
						asRowContents = Split(sRowContents & sTotals, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
				Next
				sRowContents = sRowContents & TABLE_SEPARATOR & "Total" & TABLE_SEPARATOR & FormatNumber(asTotals(0)(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(1)(0), 2, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = "<SPAN COLS=""" & (UBound(asRowContents) - 2) & """ /><B>&nbsp;</B>"
				For iIndex = 1 To 12
					If (Len(oRequest("BudgetMonth").Item) = 0) Or (InStr(1, ("," & Replace(oRequest("BudgetMonth").Item, " ", "") & ","), ("," & iIndex & ","), vbBinaryCompare) > 0) Then
						sTotals = TABLE_SEPARATOR & "<B>" & asMonthNames_es(iIndex) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(2)(iIndex), 2, True, False, True) & "</B>"
						asTotals(2)(0) = asTotals(2)(0)  + asTotals(2)(iIndex)
						sTotals = sTotals & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(3)(iIndex), 2, True, False, True) & "</B>"
						asTotals(3)(0) = asTotals(3)(0)  + asTotals(3)(iIndex)
						asRowContents = Split(sRowContents & sTotals, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
				Next
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(2)(0), 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(3)(0), 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1504 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1581(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees for the given payroll
'         group by companies and branches
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1581"
	Dim oRecordset
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCurrentID
	Dim lMinID
	Dim lMaxID
	Dim iIndex
	Dim lTotal
	Dim sOriginalRow
	Dim sTotalCount
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sErrorDescription = "No se pudieron obtener los estatus de los trámites."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.CompanyID, CompanyName, Count(Payroll_" & lPayrollID & ".EmployeeID) As EmployeesCount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Positions, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ") And (ConceptID=0) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeesHistoryList.CompanyID, CompanyName Order By EmployeesHistoryList.CompanyID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sRowContents = "&nbsp;"
			sOriginalRow = "<BRANCH_NAME />"
			sTotalCount = "<B>TOTAL</B>"
			lTotal = 0
			asCellWidths = "150"
			asCellAlignments = ""
			lMinID = CLng(oRecordset.Fields("CompanyID").Value)
			lMaxID = CLng(oRecordset.Fields("CompanyID").Value)
			Do While Not oRecordset.EOF
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
				sOriginalRow = sOriginalRow & TABLE_SEPARATOR & "<EMPLOYEES_COMPANY_" & CStr(oRecordset.Fields("CompanyID").Value) & " />"
				sTotalCount = sTotalCount & TABLE_SEPARATOR & "<B>" & FormatNumber(CLng(oRecordset.Fields("EmployeesCount").Value), 0, True, False, True) & "</B>"
				lTotal = lTotal + CLng(oRecordset.Fields("EmployeesCount").Value)
				asCellWidths = asCellWidths & ",150"
				asCellAlignments = asCellAlignments & ",LEFT"
				If CLng(oRecordset.Fields("CompanyID").Value) < lMinID Then lMinID = CLng(oRecordset.Fields("CompanyID").Value)
				If CLng(oRecordset.Fields("CompanyID").Value) > lMaxID Then lMaxID = CLng(oRecordset.Fields("CompanyID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			sRowContents = sRowContents & TABLE_SEPARATOR & "TOTAL"
			sOriginalRow = sOriginalRow & TABLE_SEPARATOR & "<TOTAL />"
			sTotalCount = sTotalCount & TABLE_SEPARATOR & "<B>" & FormatNumber(lTotal, 0, True, False, True) & "</B>"

			sErrorDescription = "No se pudieron obtener los estatus de los trámites."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.BranchID, BranchName, EmployeesHistoryList.CompanyID, CompanyName, Count(Payroll_" & lPayrollID & ".EmployeeID) As EmployeesCount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Positions, Companies, Branches Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Positions.BranchID=Branches.BranchID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ") And (ConceptID=0) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") Group By Positions.BranchID, BranchName, EmployeesHistoryList.CompanyID, CompanyName Order By Positions.BranchID, EmployeesHistoryList.CompanyID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Response.Write "<TABLE BORDER="""
						If Not bForExport Then
							Response.Write "0"
						Else
							Response.Write "1"
						End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
						asColumnsTitles = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
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

						sRowContents = sOriginalRow
						lCurrentID = -2
						lTotal = 0
						asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
						Do While Not oRecordset.EOF
							If lCurrentID <> CLng(oRecordset.Fields("BranchID").Value) Then
								If lCurrentID <> -2 Then
									For iIndex = lMinID To lMaxID
										sRowContents = Replace(sRowContents, "<EMPLOYEES_COMPANY_" & iIndex & " />", "0")
									Next
									sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(lTotal, 0, True, False, True))
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
								End If
								lCurrentID = CLng(oRecordset.Fields("BranchID").Value)
								lTotal = 0
								sRowContents = sOriginalRow
								sRowContents = Replace(sRowContents, "<BRANCH_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("BranchName").Value)))
							End If
							sRowContents = Replace(sRowContents, "<EMPLOYEES_COMPANY_" & CStr(oRecordset.Fields("CompanyID").Value) & " />", FormatNumber(CLng(oRecordset.Fields("EmployeesCount").Value), 0, True, False, True))
							lTotal = lTotal + CLng(oRecordset.Fields("EmployeesCount").Value)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						For iIndex = lMinID To lMaxID
							sRowContents = Replace(sRowContents, "<EMPLOYEES_COMPANY_" & iIndex & " />", "0")
						Next
						sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(lTotal, 0, True, False, True))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If

						asRowContents = Split(sTotalCount, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					Response.Write "</TABLE>"
				End If
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1581 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1582(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees for the given payroll
'         group by generic positions
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1582"
	Dim oRecordset
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCurrentID
	Dim lTotal
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(sCondition, "Companies.", "EmployeesHistoryList.")
	sErrorDescription = "No se pudieron obtener los estatus de los trámites."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(Payroll_" & lPayrollID & ".EmployeeID) As TotalJobs, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, GenericPositionName From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Positions, EmployeeTypes, GenericPositions Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ") And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.GenericPositionID=GenericPositions.GenericPositionID) And (ConceptID=0) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeeTypes.EmployeeTypeID, EmployeeTypeName, GenericPositionName Order By EmployeeTypeName, GenericPositionName", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("<SPAN COLS=""2"" />&nbsp;,<SPAN COLS=""2"" />Programa,<SPAN COLS=""2"" />Plazas,<SPAN COLS=""2"" />Recursos", ",", -1, vbBinaryCompare)
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

				asColumnsTitles = Split("<SPAN COLS=""2"" />Conceptos,Original,Modificado,Ocupadas,Desocupadas,Original,Modificado", ",", -1, vbBinaryCompare)
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

				lCurrentID = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				lTotal = 0
				asCellAlignments = Split(",,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeTypeID").Value) Then
						sRowContents = "<SPAN COLS=""2"" /><B>TOTAL</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(lTotal, 0, True, False, True) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lTotal = 0
						lCurrentID = CLng(oRecordset.Fields("EmployeeTypeID").Value)
					End If
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GenericPositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & 0 & TABLE_SEPARATOR & 0
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("TotalJobs").Value), 0, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & 0 & TABLE_SEPARATOR & 0 & TABLE_SEPARATOR & 0
					lTotal = lTotal + CLng(oRecordset.Fields("TotalJobs").Value)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				sRowContents = "<SPAN COLS=""2"" /><B>TOTAL</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(lTotal, 0, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>" & TABLE_SEPARATOR & "<B>0</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1582 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1583(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees for the given payroll
'         group by area and employee type
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1583"
	Dim oRecordset
	Dim sCondition
	Dim lPayrollYear
	Dim iQuarter
	Dim lCurrentID
	Dim iIndex
	Dim alTotal
	Dim sOriginalRow
	Dim sTotalCount
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "Companies.", "Employees.")
	lPayrollYear = CLng(oRequest("YearID").Item)

	sErrorDescription = "No se pudieron obtener los estatus de los trámites."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Budgets1.BudgetShortName As BudgetShortName1, Budgets2.BudgetShortName As BudgetShortName2, Budgets3.BudgetID As BudgetID3, Budgets3.BudgetShortName As BudgetShortName3, Budgets3.BudgetName As BudgetName3, Payroll_" & lPayrollYear & ".RecordDate, Sum(Payroll_" & lPayrollYear & ".ConceptAmount) As TotalAmount, Count(Payroll_" & lPayrollYear & ".EmployeeID) As EmployeesCount From Payroll_" & lPayrollYear & ", Employees, Concepts, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Payroll_" & lPayrollYear & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollYear & ".ConceptID=Concepts.ConceptID) And (Concepts.BudgetID=Budgets3.BudgetID) And (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (Concepts.ConceptID In (4,5,6,12,16,17,20,21,24,30,32,35,36,37,45,46,92,122,125)) And (Concepts.StartDate<=" & lPayrollYear & "0101) And (Concepts.EndDate>=" & lPayrollYear & "1231) " & sCondition & " Group By Budgets1.BudgetShortName, Budgets2.BudgetShortName, Budgets3.BudgetID, Budgets3.BudgetShortName, Budgets3.BudgetName, Payroll_" & lPayrollYear & ".RecordDate Order By Budgets1.BudgetShortName, Budgets2.BudgetShortName, Budgets3.BudgetShortName, Payroll_" & lPayrollYear & ".RecordDate", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sOriginalRow = "<BUDGET1_SHORT_NAME />" & TABLE_SEPARATOR & "<BUDGET2_SHORT_NAME />" & TABLE_SEPARATOR & "<BUDGET3_SHORT_NAME />" & TABLE_SEPARATOR & "<BUDGET3_NAME />" & TABLE_SEPARATOR & "<QUARTER_0 />" & TABLE_SEPARATOR & "<QUARTER_1 />" & TABLE_SEPARATOR & "<QUARTER_2 />" & TABLE_SEPARATOR & "<QUARTER_3 />" & TABLE_SEPARATOR & "<TOTAL />" & TABLE_SEPARATOR & "<EMPLOYEES />"
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asCellWidths = Split("50,50,50,200,150,150,150,150,150,150", ",", -1, vbBinaryCompare)
				asColumnsTitles = Split("<SPAN COLS=""4"" />&nbsp;,<SPAN COLS=""5"" />ISSSTE-ASEGURADOR,ISSSTE-ASEGURADOR", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asColumnsTitles = Split("Partida,Subpartida,Tipo de pago,Denominación,Ejercicio<BR />Enero-Marzo,Ejercicio<BR />Abril-Junio,Ejercicio<BR />Julio-Septiembre,Ejercicio<BR />Octubre-Diciembre,Enero-Diciembre,Casos", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				sRowContents = sOriginalRow
				lCurrentID = -2
				alTotal = Split("0,0,0,0,0,0;0,0,0,0,0,0", ";")
				alTotal(0) = Split(alTotal(0), ",")
				alTotal(1) = Split(alTotal(1), ",")
				For iIndex = 0 To UBound(alTotal(0))
					alTotal(0)(iIndex) = 0
					alTotal(1)(iIndex) = 0
				Next
				asCellAlignments = Split("CENTER,CENTER,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("BudgetID3").Value) Then
						If lCurrentID <> -2 Then
							For iIndex = 0 To UBound(alTotal(0)) - 2
								sRowContents = Replace(sRowContents, "<QUARTER_" & iIndex & " />", FormatNumber(alTotal(0)(iIndex), 2, True, False, True))
								alTotal(1)(iIndex) = alTotal(1)(iIndex) + alTotal(0)(iIndex)
								alTotal(0)(iIndex) = 0
							Next
							sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(alTotal(0)(4), 2, True, False, True))
							alTotal(1)(4) = alTotal(1)(4) + alTotal(0)(4)
							alTotal(0)(4) = 0
							sRowContents = Replace(sRowContents, "<EMPLOYEES />", FormatNumber(alTotal(0)(5), 0, True, False, True))
							alTotal(1)(5) = alTotal(1)(5) + alTotal(0)(5)
							alTotal(0)(5) = 0
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
						lCurrentID = CLng(oRecordset.Fields("BudgetID3").Value)
						sRowContents = sOriginalRow
						sRowContents = Replace(sRowContents, "<BUDGET1_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName1").Value)))
						sRowContents = Replace(sRowContents, "<BUDGET2_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName2").Value)))
						sRowContents = Replace(sRowContents, "<BUDGET3_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName3").Value)))
						sRowContents = Replace(sRowContents, "<BUDGET3_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("BudgetName3").Value)))
					End If
					iQuarter = CInt(Right(CStr(oRecordset.Fields("RecordDate").Value), Len("0000")))
					If iQuarter < 400 Then
						alTotal(0)(0) = alTotal(0)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					ElseIf iQuarter < 700 Then
						alTotal(0)(1) = alTotal(0)(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
					ElseIf iQuarter < 1000 Then
						alTotal(0)(2) = alTotal(0)(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						alTotal(0)(3) = alTotal(0)(3) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					alTotal(0)(4) = alTotal(0)(4) + CDbl(oRecordset.Fields("TotalAmount").Value)
					alTotal(0)(5) = alTotal(0)(5) + CDbl(oRecordset.Fields("EmployeesCount").Value)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				For iIndex = 0 To UBound(alTotal(0)) - 2
					sRowContents = Replace(sRowContents, "<QUARTER_" & iIndex & " />", FormatNumber(alTotal(0)(iIndex), 2, True, False, True))
					alTotal(1)(iIndex) = alTotal(1)(iIndex) + alTotal(0)(iIndex)
				Next
				sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(alTotal(0)(4), 2, True, False, True))
				alTotal(1)(4) = alTotal(1)(4) + alTotal(0)(4)
				sRowContents = Replace(sRowContents, "<EMPLOYEES />", FormatNumber(alTotal(0)(5), 0, True, False, True))
				alTotal(1)(5) = alTotal(1)(5) + alTotal(0)(5)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = sOriginalRow
				sRowContents = Replace(sRowContents, "<BUDGET1_SHORT_NAME />", "&nbsp;")
				sRowContents = Replace(sRowContents, "<BUDGET2_SHORT_NAME />", "&nbsp;")
				sRowContents = Replace(sRowContents, "<BUDGET3_SHORT_NAME />", "&nbsp;")
				sRowContents = Replace(sRowContents, "<BUDGET3_NAME />", "<B>TOTAL CAPÍTULO 1000</B>")
				For iIndex = 0 To UBound(alTotal(1)) - 2
					sRowContents = Replace(sRowContents, "<QUARTER_" & iIndex & " />", "<B>" & FormatNumber(alTotal(1)(iIndex), 2, True, False, True) & "</B>")
				Next
				sRowContents = Replace(sRowContents, "<TOTAL />", "<B>" & FormatNumber(alTotal(1)(4), 2, True, False, True) & "</B>")
				sRowContents = Replace(sRowContents, "<EMPLOYEES />", "<B>" & FormatNumber(alTotal(1)(5), 0, True, False, True) & "</B>")
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1583 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1584(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees for the given payroll
'         group by area and employee type
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1584"
	Dim oRecordset
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCurrentID
	Dim iIndex
	Dim alTotal
	Dim sOriginalRow
	Dim sTotalCount
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(sCondition, "Companies.", "EmployeesHistoryList.")

	sErrorDescription = "No se pudieron obtener los estatus de los trámites."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionTypes.PositionTypeID, PositionTypeName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, ParentAreas.AreaID, ParentAreas.AreaCode, ParentAreas.AreaName, Count(Payroll_" & lPayrollID & ".EmployeeID) As EmployeesCount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Positions, PositionTypes, EmployeeTypes, Areas, Areas As ParentAreas Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ") And (ConceptID=0) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (ParentAreas.StartDate<=" & lForPayrollID & ") And (ParentAreas.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By PositionTypes.PositionTypeID, PositionTypeName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, ParentAreas.AreaID, ParentAreas.AreaCode, ParentAreas.AreaName Order By ParentAreas.AreaCode", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sOriginalRow = "<AREA_CODE />" & TABLE_SEPARATOR & "<AREA_NAME />" & TABLE_SEPARATOR & "<TYPE_0 />" & TABLE_SEPARATOR & "<TYPE_2 />" & TABLE_SEPARATOR & "<TYPE_1 />" & TABLE_SEPARATOR & "<TYPE_3 />" & TABLE_SEPARATOR & "<TYPE_4 />" & TABLE_SEPARATOR & "<TYPE_5 />" & TABLE_SEPARATOR & "<TOTAL />"
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asCellWidths = Split("100,150,150,150,150,150,150,150,150", ",", -1, vbBinaryCompare)
				asColumnsTitles = Split("<SPAN COLS=""2"" />&nbsp;,<SPAN COLS=""3"" />No de plazas ocupadas: Confianza y Base,<SPAN COLS=""3"" />Otro tipo de personal,&nbsp;", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asColumnsTitles = Split("Código,Entidad,Empleados,Funcionarios,Base,Honorarios,Becarios,Residentes,TOTAL", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				sRowContents = sOriginalRow
				lCurrentID = -2
				alTotal = Split("0,0,0,0,0,0,0;0,0,0,0,0,0,0", ";")
				alTotal(0) = Split(alTotal(0), ",")
				alTotal(1) = Split(alTotal(1), ",")
				For iIndex = 0 To UBound(alTotal(0))
					alTotal(0)(iIndex) = 0
					alTotal(1)(iIndex) = 0
				Next
				asCellAlignments = Split("CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("AreaID").Value) Then
						If lCurrentID <> -2 Then
							For iIndex = 0 To UBound(alTotal(0)) - 1
								sRowContents = Replace(sRowContents, "<TYPE_" & iIndex & " />", FormatNumber(alTotal(0)(iIndex), 0, True, False, True))
								alTotal(1)(iIndex) = alTotal(1)(iIndex) + alTotal(0)(iIndex)
								alTotal(0)(6) = alTotal(0)(6) + alTotal(0)(iIndex)
								alTotal(0)(iIndex) = 0
							Next
							sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(alTotal(0)(6), 0, True, False, True))
							alTotal(1)(6) = alTotal(1)(6) + alTotal(0)(6)
							alTotal(0)(6) = 0
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
						lCurrentID = CLng(oRecordset.Fields("AreaID").Value)
						sRowContents = sOriginalRow
						sRowContents = Replace(sRowContents, "<AREA_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)))
						sRowContents = Replace(sRowContents, "<AREA_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)))
					End If
					If (CLng(oRecordset.Fields("PositionTypeID").Value) = 2) And (CLng(oRecordset.Fields("EmployeeTypeID").Value) <> 1) Then
						alTotal(0)(0) = CLng(oRecordset.Fields("EmployeesCount").Value)
					Else
						alTotal(0)(CLng(oRecordset.Fields("PositionTypeID").Value)) = CLng(oRecordset.Fields("EmployeesCount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop

				For iIndex = 0 To UBound(alTotal(0)) - 1
					sRowContents = Replace(sRowContents, "<TYPE_" & iIndex & " />", FormatNumber(alTotal(0)(iIndex), 0, True, False, True))
					alTotal(1)(iIndex) = alTotal(1)(iIndex) + alTotal(0)(iIndex)
					alTotal(0)(6) = alTotal(0)(6) + alTotal(0)(iIndex)
					alTotal(0)(iIndex) = 0
				Next
				sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber(alTotal(0)(6), 0, True, False, True))
				alTotal(1)(6) = alTotal(1)(6) + alTotal(0)(6)
				alTotal(0)(6) = 0
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = sOriginalRow
				sRowContents = Replace(sRowContents, "<AREA_CODE />", "&nbsp;")
				sRowContents = Replace(sRowContents, "<AREA_NAME />", "<B>TOTAL</B>")
				For iIndex = 0 To UBound(alTotal(1)) - 1
					sRowContents = Replace(sRowContents, "<TYPE_" & iIndex & " />", "<B>" & FormatNumber(alTotal(1)(iIndex), 0, True, False, True) & "</B>")
				Next
				sRowContents = Replace(sRowContents, "<TOTAL />", "<B>" & FormatNumber(alTotal(1)(6), 0, True, False, True) & "</B>")
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1584 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1603(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paperworks status
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1603"
	Dim oRecordset
	Dim sCondition
	Dim sFontBegin
	Dim sFontEnd
	Dim oPpwkStartDate
	Dim oEstimatedDate
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "(StartDate", "(Paperworks.StartDate"), "(EndDate", "(Paperworks.EndDate")
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los estatus de los trámites."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkTypeName, StatusName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwnersLKP.EndDate As EndDate2 From Paperworks, PaperworkTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, StatusName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwnersLKP.EndDate As EndDate2 From Paperworks, PaperworkTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("No. de trámite,No. documento,Tipo de trámite,Responsable,Estatus,Fecha de recepción,Fecha límite de respuesta,Fecha de atención", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,200,200,100,200,200,200", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				asCellAlignments = Split(",,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					oPpwkStartDate = GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value))
					oEstimatedDate = GetDateFromSerialNumber(CStr(oRecordset.Fields("EstimatedDate").Value))
					If (CLng(oRecordset.Fields("EndDate").Value) = 0) And (CLng(oRecordset.Fields("EstimatedDate").Value) < CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>"
						sFontEnd = "</B></FONT>"
					ElseIf CLng(oRecordset.Fields("EndDate").Value) > CLng(oRecordset.Fields("EstimatedDate").Value) Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					ElseIf (CLng(oRecordset.Fields("EndDate").Value) = 0) And (CLng(oRecordset.Fields("EstimatedDate").Value) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))) Then
						sFontBegin = "<FONT COLOR=""#D2D200""><B>"
						sFontEnd = "</B></FONT>"
					ElseIf (CLng(oRecordset.Fields("EndDate").Value) = 0) And DateDiff("d", oPpwkStartDate, oEstimatedDate) > (DateDiff("d", Date(), oEstimatedDate) * 2) Then
						sFontBegin = "<FONT COLOR=""#D2D200"">"
						sFontEnd = "</FONT>"
					ElseIf (CLng(oRecordset.Fields("EndDate").Value) > 0) Then
						sFontBegin = "<FONT COLOR=""#00D200"">"
						sFontEnd = "</FONT>"
					End If
					sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value)) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value)) & sFontEnd
					If CLng(oRecordset.Fields("EndDate2").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "En trámite" & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "Cerrado" & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1) & sFontEnd
					If CLng(oRecordset.Fields("EndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & "<CENTER>---</CENTER>" & sFontEnd
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1603 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1604(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paperworks status
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1604"
	Dim oRecordset
	Dim sCondition
	Dim asConditions
	Dim asTitles
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	asConditions = Split(", And (EstimatedDate<=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EndDate=0), And (EstimatedDate<EndDate) And (EstimatedDate>0) And (EndDate>0), And ((EstimatedDate>=EndDate) Or (EstimatedDate=0)) And (EndDate>0)", ",")
	asTitles = Split(",ASUNTOS ABIERTOS Y DESFASADOS,ASUNTOS CERRADOS PERO DESFASADOS,ASUNTOS RESUELTOS (CERRADOS A TIEMPO)", ",")
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "(StartDate", "(Paperworks.StartDate"), "(EndDate", "(Paperworks.EndDate")
	oStartDate = Now()

	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	sFilePath = Server.MapPath(sFileName & ".xls")
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()

	For iIndex = 1 To UBound(asConditions)
		If InStr(1, ("," & Replace(oRequest("Include").Item, " ", "") & ","), ("," & iIndex & ","), vbBinaryCompare) > 0 Then
			sErrorDescription = "No se pudieron obtener los estatus de los trámites."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkTypeName, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, StatusName From Paperworks, PaperworkTypes, StatusPaperworks Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) " & sCondition & asConditions(iIndex) & " Order By PaperworkNumber", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, StatusName From Paperworks, PaperworkTypes, StatusPaperworks Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) " & sCondition & asConditions(iIndex) & " Order By PaperworkNumber -->" & vbNewLine
			If lErrorNumber = 0 Then
				lErrorNumber = AppendTextToFile(sFilePath, "<B>" & asTitles(iIndex) & "</B><BR /><BR />", sErrorDescription)
				If Not oRecordset.EOF Then
					lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine, sErrorDescription)
						asColumnsTitles = Split("No. de trámite,No. documento,Tipo de trámite,Fecha de recepción,Fecha límite de respuesta,Fecha de atención,Días transcurridos", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,100,200,200,200,200,100", ",", -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

						asCellAlignments = Split(",,,,,,,RIGHT", ",", -1, vbBinaryCompare)
						Do While Not oRecordset.EOF
							sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
							If CLng(oRecordset.Fields("EndDate").Value) > 0 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
								sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value))) + 1)
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
								sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), Date()) + 1)
							End If

							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
					lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)
				Else
					lErrorNumber = AppendTextToFile(sFilePath, "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B><BR /><BR />", sErrorDescription)
				End If
			End If
		End If
	Next
	lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
	If lErrorNumber = 0 Then
		Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
		sErrorDescription = "No se pudieron guardar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
	End If
	oEndDate = Now()
	If (lErrorNumber = 0) And B_USE_SMTP Then
		If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1604 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1605(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To generate a word document of employees that have syndicate license
'         Departamento técnico
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1605"
	Dim oRecordset
	Dim sCondition
	Dim lErrorNumber
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sRowContents
	Dim asRowContents
	Dim sContents
	Dim asContents
	Dim asCellAlignments
	Dim asCellWidths
	Dim sDocumentTemplate
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sTextLine
	Dim iLicenseSyndicateTypeID

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(sCondition) > 0 Then 
		sCondition = Replace(sCondition, "XXX", "DocumentsForLicenses.DocumentLicense")
	End If
	
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sErrorDescription = "No se pudo obtener la información del empleado."
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	sDocumentName = sFilePath & "LicenciasSindicales_" & sDate & ".doc"
	
	If lErrorNumber = 0 Then
		sContents = GetFileContents(Server.MapPath("Templates\1. SNTISSSTE.htm"), sErrorDescription)
		If Len(sContents)> 0 Then
			asContents = Split(sContents, vbNewLine, -1, vbBinaryCompare)
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			sRowContents = RTF_BEGIN_V & " "
			sRowContents = sRowContents & RTF_DEFAULT_TITLE
			sRowContents = sRowContents & RTF_HEADER_BEGIN
			sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN
			lErrorNumber = SaveTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "{ " & GetFileContents(Server.MapPath("Templates\LogoISSSTE_RTF.txt"), sErrorDescription)
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & RTF_FONT15_START & " " & asContents(0) & RTF_FONT_END & " "
			sRowContents = sRowContents & TABLE_SEPARATOR
			sRowContents = sRowContents & GetFileContents(Server.MapPath("Templates\LogoEscudoNacional_RTF.txt"), sErrorDescription)
			sRowContents = sRowContents & TABLE_SEPARATOR
			asCellAlignments = Split("LEFT,CENTER,RIGHT", ",", -1, vbBinaryCompare)
			asCellWidths = Split("3000,8000,10000", ",", -1, vbBinaryCompare)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, False, sDocumentName, sErrorDescription)
			sRowContents = "}"
			sRowContents = sRowContents & RTF_PARAGRAPH_END
			sRowContents = sRowContents & RTF_HEADER_END
			sRowContents = sRowContents & RTF_FOOTER_BEGIN & " " & RTF_PARAGRAPH_BEGIN & RTF_CENTER
			sRowContents = sRowContents & RTF_FONT15_START & " " & asContents(22) & RTF_FONT_END & RFT_NEW_LINE
			sRowContents = sRowContents & RTF_FONT15_START & " " & asContents(23) & RTF_FONT_END
			sRowContents = sRowContents & RTF_PARAGRAPH_END
			sRowContents = sRowContents & RTF_FOOTER_END
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, DocumentsForLicenses.DocumentTemplate, LicenseSyndicateTypes.LicenseSyndicateTypeName, DocumentsForLicenses.DocumentForLicenseNumber, DocumentForCancelLicenseNumber, DocumentsForLicenses.RequestNumber, DocumentsForLicenses.DocumentLicenseDate, DocumentsForLicenses.LicenseStartDate, DocumentsForLicenses.LicenseEndDate, LicenseCancelDate, PositionShortName, Positions.PositionName, LevelShortName, Areas.AreaCode, Areas.AreaName As Area, ParentAreas.AreaName As ParentArea From Areas, Areas As ParentAreas, Employees, DocumentsForLicenses, LicenseSyndicateTypes, Jobs, Levels, Positions Where (Employees.EmployeeID=DocumentsForLicenses.EmployeeID) And (Employees.JobID=Jobs.JobID)And (Jobs.AreaID=Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.LevelID=Levels.LevelID) And (DocumentsForLicenses.LicenseSyndicateTypeID = LicenseSyndicateTypes.LicenseSyndicateTypeID)" & sCondition, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, DocumentsForLicenses.DocumentTemplate, LicenseSyndicateTypes.LicenseSyndicateTypeName, DocumentsForLicenses.DocumentForLicenseNumber, DocumentForCancelLicenseNumber, DocumentsForLicenses.RequestNumber, DocumentsForLicenses.DocumentLicenseDate, DocumentsForLicenses.LicenseStartDate, DocumentsForLicenses.LicenseEndDate, LicenseCancelDate, PositionShortName, Positions.PositionName, LevelShortName, Areas.AreaCode, Areas.AreaName As Area, ParentAreas.AreaName As ParentArea From Areas, Areas As ParentAreas, Employees, DocumentsForLicenses, LicenseSyndicateTypes, Jobs, Levels, Positions Where (Employees.EmployeeID=DocumentsForLicenses.EmployeeID) And (Employees.JobID=Jobs.JobID)And (Jobs.AreaID=Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.LevelID=Levels.LevelID) And (DocumentsForLicenses.LicenseSyndicateTypeID = LicenseSyndicateTypes.LicenseSyndicateTypeID)" & sCondition & " -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						sDocumentTemplate=CStr(oRecordset.Fields("DocumentTemplate").Value)
						sContents = GetFileContents(Server.MapPath("Templates\" & sDocumentTemplate), sErrorDescription)
						If Len(sContents)> 0 Then
							asContents = Split(sContents, vbNewLine, -1, vbBinaryCompare)
							sRowContents = RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(1) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(2) & RTF_FONT_END & RTF_PARAGRAPH_END
							Select Case sDocumentTemplate
								Case "1. SNTISSSTE.htm", "2. FSTSE.htm"
									sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(3) & CStr(oRecordset.Fields("DocumentForLicenseNumber").Value) & "/" & Year(Date()) & RTF_FONT_END & RTF_PARAGRAPH_END
									iLicenseSyndicateTypeID = 1
								Case "3. Cancela SNTISSSTE.htm" , "4. Cancela FSTSE.htm"
									sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(3) & CStr(oRecordset.Fields("DocumentForCancelLicenseNumber").Value) & "/" & Year(Date()) & RTF_FONT_END & RTF_PARAGRAPH_END
									iLicenseSyndicateTypeID = 2
							End Select	
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_RIGHT & RTF_FONT20_START & " " & " México, D.F. a " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("DocumentLicenseDate").Value), -1, -1, -1) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(4) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(5) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(6) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & asContents(7) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RFT_NEW_LINE
							sTextLine = Replace(Replace(Replace(Replace(Replace(asContents(8), "<REQUEST_NUMBER />", CStr(oRecordset.Fields("RequestNumber").Value)),"<SHORT_DATE />", Right(CStr(Year(Date)),Len("00"))), "<LICENSE_CANCEL_DATE />", DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseCancelDate").Value))),"<DOCUMENT_FOR_LICENSE_NUMBER>", CStr(oRecordset.Fields("DocumentForLicenseNumber").Value)), "<YEAR />", Year(Date))
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_JUSTIFIED & RTF_FONT20_START & " " & sTextLine & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RFT_NEW_LINE
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_CENTER & RTF_FONT17_START & RTF_BOLD & " " & asContents(9) & RTF_FONT_END & RTF_PARAGRAPH_END
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							asCellWidths = Split("4000,9900", ",", -1, vbBinaryCompare)
							asCellAlignments = Split("LEFT,LEFT", ",", -1, vbBinaryCompare)
							sRowContents = "{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Nombre:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)) & RTF_FONT_END
							Else
								sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)) & RTF_FONT_END
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							If iLicenseSyndicateTypeID = 1 Then
								sRowContents = RTF_FONT19_START & "Tipo de Licencia:" & RTF_FONT_END
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("LicenseSyndicateTypeName").Value)) & RTF_FONT_END
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
								sRowContents = "}{"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

								sRowContents = RTF_FONT19_START & "Vigencia:" & RTF_FONT_END
								sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & RTF_FONT19_START & " " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseStartDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseEndDate").Value)) & RTF_FONT_END
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
								sRowContents = "}{"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							End If

							sRowContents = RTF_FONT19_START & "Código de Puesto, nivel y subnivel:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "   " & Left(CStr(oRecordset.Fields("LevelShortName").Value),Len("00"))& " " & Right(CStr(oRecordset.Fields("LevelShortName").Value),Len("0")) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Denominación del puesto:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Registro Federal Contribuyentes:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Número de empleado:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CStr(oRecordset.Fields("EmployeeNumber").Value) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Adscripción presupuestal:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("Area").Value)) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("ParentArea").Value)) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "Clave de adscripción:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}{"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_FONT19_START & "CURP:" & RTF_FONT_END
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & RTF_FONT19_START & " " & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & RTF_FONT_END
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, True, sDocumentName, sErrorDescription)
							sRowContents = "}"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

							sRowContents = RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT18_START & RTF_BOLD & " " & "" & RTF_FONT_END & RTF_PARAGRAPH_END
							If Len(asContents(10)) > 0 Then
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_JUSTIFIED & RTF_FONT20_START & " " & asContents(10) & RTF_FONT_END & RTF_PARAGRAPH_END
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							End If
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & " " & asContents(11) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_CENTER & RTF_FONT20_START & RTF_BOLD & " " & asContents(12) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_CENTER & RTF_FONT20_START & RTF_BOLD & " " & asContents(13) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT20_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							If iLicenseSyndicateTypeID = 2 Then
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							End If
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_CENTER & RTF_FONT20_START & RTF_BOLD & " " & asContents(14) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							If iLicenseSyndicateTypeID = 2 Then
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							End If
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(15) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(16) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(17) & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(18) & RTF_FONT_END & RTF_PARAGRAPH_END
							If Len(asContents(19))>0 Then
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(19) & RTF_FONT_END & RTF_PARAGRAPH_END
							End If
							If Len(asContents(20))>0 Then
								sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & " " & asContents(20) & RTF_FONT_END & RTF_PARAGRAPH_END
							End If
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & "        " & CStr(oRecordset.Fields("ParentArea").Value) &  ".- Para su atención procedente.- Presente" & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & RTF_FONT_END & RTF_PARAGRAPH_END
							sRowContents = sRowContents & RTF_PARAGRAPH_BEGIN & RTF_LEFT & RTF_FONT13_START & RTF_BOLD & " " & asContents(21) & RTF_FONT_END & RTF_PARAGRAPH_END & RFT_NEW_PAGE
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
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
				Else
					sRowContents = RTF_END
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
					End If
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
					oRecordset.Close
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1605 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1606(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display sindicate licenses documents with employee's information
'         Departamento técnico
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1606"
	Dim sCondition
	Dim lPayrollID
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

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(sCondition) > 0 Then 
		sCondition = Replace(sCondition, "XXX", "DocumentsForLicenses.DocumentLicense")
	End If

    oStartDate = Now()
	sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, DocumentsForLicenses.DocumentTemplate, LicenseSyndicateTypes.LicenseSyndicateTypeName, DocumentsForLicenses.DocumentForLicenseNumber, DocumentForCancelLicenseNumber, DocumentsForLicenses.RequestNumber, DocumentsForLicenses.DocumentLicenseDate, DocumentsForLicenses.LicenseStartDate, DocumentsForLicenses.LicenseEndDate, LicenseCancelDate, PositionShortName, Positions.PositionName, LevelShortName, Areas.AreaCode, Areas.AreaName As Area, ParentAreas.AreaName As ParentArea From Areas, Areas As ParentAreas, Employees, DocumentsForLicenses, LicenseSyndicateTypes, Jobs, Levels, Positions Where (Employees.EmployeeID=DocumentsForLicenses.EmployeeID) And (Employees.JobID=Jobs.JobID)And (Jobs.AreaID=Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.LevelID=Levels.LevelID) And (DocumentsForLicenses.LicenseSyndicateTypeID = LicenseSyndicateTypes.LicenseSyndicateTypeID)" & sCondition, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, DocumentsForLicenses.DocumentTemplate, LicenseSyndicateTypes.LicenseSyndicateTypeName, DocumentsForLicenses.DocumentForLicenseNumber, DocumentForCancelLicenseNumber, DocumentsForLicenses.RequestNumber, DocumentsForLicenses.DocumentLicenseDate, DocumentsForLicenses.LicenseStartDate, DocumentsForLicenses.LicenseEndDate, LicenseCancelDate, PositionShortName, Positions.PositionName, LevelShortName, Areas.AreaCode, Areas.AreaName As Area, ParentAreas.AreaName As ParentArea From Areas, Areas As ParentAreas, Employees, DocumentsForLicenses, LicenseSyndicateTypes, Jobs, Levels, Positions Where (Employees.EmployeeID=DocumentsForLicenses.EmployeeID) And (Employees.JobID=Jobs.JobID)And (Jobs.AreaID=Areas.AreaID) And (Areas.ParentID = ParentAreas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.LevelID=Levels.LevelID) And (DocumentsForLicenses.LicenseSyndicateTypeID = LicenseSyndicateTypes.LicenseSyndicateTypeID)" & sCondition & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sDocumentName = sFilePath & "EmpLicSindical_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = SaveTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TR>"
					sRowContents = sRowContents & "<TD>No. oficio</TD>"
					sRowContents = sRowContents & "<TD>No. oficio cancela</TD>"
					sRowContents = sRowContents & "<TD>Fecha documento</TD>"
					sRowContents = sRowContents & "<TD>Plantilla</TD>"
					sRowContents = sRowContents & "<TD>No. de solicitud</TD>"
					sRowContents = sRowContents & "<TD>Tipo de licencia</TD>"
					sRowContents = sRowContents & "<TD>Fecha inicio licencia</TD>"
					sRowContents = sRowContents & "<TD>Fecha término licencia</TD>"
					sRowContents = sRowContents & "<TD>Fecha cancela licencia</TD>"
					sRowContents = sRowContents & "<TD>No. Empleado</TD>"
					sRowContents = sRowContents & "<TD>Nombre</TD>"
					sRowContents = sRowContents & "<TD>Puesto</TD>"
					sRowContents = sRowContents & "<TD>Nivel y subnivel</TD>"
					sRowContents = sRowContents & "<TD>Denominacion del puesto</TD>"
					sRowContents = sRowContents & "<TD>RFC</TD>"
					sRowContents = sRowContents & "<TD>CURP</TD>"
					sRowContents = sRowContents & "<TD>Adscripción presupuestal</TD>"
					sRowContents = sRowContents & "<TD>Delegación</TD>"
					sRowContents = sRowContents & "<TD>Clave de adscripción</TD>"
				sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("DocumentForLicenseNumber").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("DocumentForCancelLicenseNumber").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("DocumentLicenseDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("DocumentTemplate").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("RequestNumber").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("LicenseSyndicateTypeName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseStartDate").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseEndDate").Value)) & "</TD>"
						If CLng(oRecordset.Fields("LicenseCancelDate").Value) <> 0 Then
							sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("LicenseCancelDate").Value)) & "</TD>"
						Else
							sRowContents = sRowContents & "<TD></TD>"
						End If	
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) &  "</TD>"
						Else
							sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) &  "</TD>"
						End If
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PositionShortName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("LevelShortName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PositionName").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("RFC").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("CURP").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("Area").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ParentArea").Value) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("AreaCode").Value) & "</TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				'If lErrorNumber = 0 Then
				'	Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				'	sErrorDescription = "No se pudieron guardar la información del reporte."
				'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				'End If
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

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1606 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1607(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1607"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(oRequest("YearID").Item) > 0 Then sCondition = sCondition & " And (Paperworks.StartDate>=" & oRequest("YearID").Item & "0000) And (Paperworks.StartDate<=" & oRequest("YearID").Item & "9999)"
	If Len(oRequest("Closed").Item) > 0 Then
		If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then
			sCondition = sCondition & " And (Paperworks.EndDate=0)"
		Else
			sCondition = sCondition & " And (Paperworks.EndDate>0)"
		End If
	End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkTypeName, SubjectTypeName, StatusName, PriorityName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, Owners.EmployeeName As OwnerName1, Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName2 As OwnerLastName2, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, Areas.AreaID, Areas.AreaName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, Priorities, PaperworkOwnersLKP, PaperworkOwners, Employees As Owners, PaperworkActions, PaperworkSenders, Areas Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.EmployeeID=Owners.EmployeeID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkSenders.AreaID=Areas.AreaID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners.LevelID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PriorityName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, Owners.EmployeeName As OwnerName1, Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName2 As OwnerLastName2, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, Areas.AreaID, Areas.AreaName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, Priorities, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders, Areas Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkSenders.AreaID=Areas.AreaID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners.LevelID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<B>Fecha y  hora de la generación del reporte: " & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), Hour(Time()), Minute(Time()), Second(Time())) & "</B><BR /><BR />", sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("C.C.;;;Fecha del documento;;;Fecha límite;;;Documento de origen;;;Fecha de recepción;;;Fecha de descargo;;;Prioridad;;;Fecha que se turnó;;;Oficio de descargo;;;Remitente;;;Estatus;;;Procedencia;;;Responsable;;;Tipo de asunto;;;Asunto;;;Observaciones;;;Tipo de trámite;;;Acciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				Do While Not oRecordset.EOF
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerEndDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "---"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("ClosingNumber").Value)) & """)"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "---"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value))
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "En trámite"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value))
						If CLng(oRecordset.Fields("AreaID").Value) > -1 Then sRowContents = sRowContents & CleanStringForHTML(" (" & CStr(oRecordset.Fields("AreaName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerAreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkActionName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1607 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1608(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1608"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
    Dim lStartDate
	Dim oEndDate
    Dim bVencido
	Dim lErrorNumber
    Dim sTables
    Dim bFromRequest
    Dim sID
    Dim iConta

	sID = oRequest("UserID").Item
	bFromRequest = (Err.number = 0)
	Err.clear

    Dim sFolio
    Dim sFechaDoc
    Dim sResponsable
    Dim sTipoTramite
    Dim sEstatus
    Dim sTipoAsunto
    Dim sPrioridad
    Dim sFechaLimite
    'NUMERO DE FOLIO
    If (Len(oRequest("FilterStartNumber").Item) > 0 And Len(oRequest("FilterEndNumber").Item) > 0) Then
        sFolio = CStr(oRequest("FilterStartNumber").Item) & " - " & CStr(oRequest("FilterEndNumber").Item)
    ElseIf (Len(oRequest("FilterStartNumber").Item) > 0 And Len(oRequest("FilterEndNumber").Item) = 0) Then
        sFolio = CStr(oRequest("FilterStartNumber").Item) & " - ?"
    ElseIf (Len(oRequest("FilterStartNumber").Item) = 0 And Len(oRequest("FilterEndNumber").Item) > 0) Then
        sFolio = "? - " & CStr(oRequest("FilterEndNumber").Item)
    Else
        sFolio = "-Todos"
    End If
    'FECHA DEL DOCUMENTO
	lErrorNumber = GetDateRank(oRequest, "PaperworkStartStart", "PaperworkStartEnd", True, sFechaDoc)
	'If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de recepción:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
    'TIPO DE TRAMITE
    If bFromRequest Then
	    sID = oRequest("PaperworkTypeID").Item
    Else
	    sID = GetParameterFromURLString(oRequest, "PaperworkTypeID")
    End If
    If Len(sID) > 0 Then
	    Call GetNameFromTable(oADODBConnection, "PaperworkTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sTipoTramite, "")
    Else
	    sTipoAsunto = "-Todos"
    End If
    'RESPONSABLE
	If bFromRequest Then
		sID = oRequest("OwnerIDs").Item
	Else
		sID = GetParameterFromURLString(oRequest, "OwnerIDs")
	End If
	If Len(sID) > 0 Then
		Call GetNameFromTable(oADODBConnection, "PaperworkOwners", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sResponsable, "")
	Else
		sResponsable = "-Todos"
	End If
    'ESTATUS
    If bFromRequest Then
	    sID = oRequest("PaperworkStatusID").Item
    Else
	    sID = GetParameterFromURLString(oRequest, "PaperworkStatusID")
    End If
    If Len(sID) > 0 Then
	    Call GetNameFromTable(oADODBConnection, "StatusPaperworks", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sEstatus, "")
    Else
	    sEstatus = "-Todos"
    End If
    'TIPO DE ASUNTO
	If bFromRequest Then
		sID = oRequest("SubjectTypeID").Item
	Else
		sID = GetParameterFromURLString(oRequest, "SubjectTypeID")
	End If
	If Len(sID) > 0 Then
		Call GetNameFromTable(oADODBConnection, "SubjectTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sTipoAsunto, "")
	Else
		sTipoAsunto = "-Todos"
	End If
    'PRIORIDAD
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_PRIORITY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PriorityID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PriorityID")
		End If
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Priorities", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sPrioridad, "")
		Else
			sPrioridad = "-Todos"
		End If
	End If
    'FECHA LIMITE
    lErrorNumber = GetDateRank(oRequest, "PaperworkEstimatedStart", "PaperworkEstimatedEnd", True, sFechaLimite)

	sTables = ""
	If InStr(1, sCondition, "Jobs", vbBinaryCompare) > 0 Then sTables = sTables & ", Jobs, Areas"
	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	oStartDate = Now()
    lStartDate = CLng(Left(GetSerialNumberForDate(oStartDate), Len("00000000")))
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName, PriorityName, SubjectTypeName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks, Priorities, SubjectTypes" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) " & sCondition & " Order By PaperworkNumber", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName, PriorityName, SubjectTypeName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks, Priorities, SubjectTypes" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) " & sCondition & " Order By PaperworkNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
            iConta = 0
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			asCellWidths = Split("80,200,150,150,100,100,100,100,250", ",", -1, vbBinaryCompare)

			lErrorNumber = AppendTextToFile(sFilePath, "<B>Fecha y  hora de la generación del reporte: " & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), Hour(Time()), Minute(Time()), Second(Time())) & "</B><BR /><BR />", sErrorDescription)
            lErrorNumber = AppendTextToFile(sFilePath, "<FONT COLOR=""#FF0000""><B>Condiciones de selección del reporte:</B></FONT>", sErrorDescription)
            lErrorNumber = AppendTextToFile(sFilePath, "<BR />", sErrorDescription)

            lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
                asColumnsTitles = Split("=T(""" & sFolio & """);;;" & sResponsable & ";;;" & sFechaDoc & ";;;" & sFechaLimite & ";;;Ven;;;" & sTipoAsunto & ";;;" & sPrioridad & ";;;" & sEstatus & ";;;Asunto", LIST_SEPARATOR, -1, vbBinaryCompare)
                lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainTextWidth(asColumnsTitles, asCellWidths, True, 1, sErrorDescription), sErrorDescription)

			asColumnsTitles = Split(";;;;;;;;;;;;;;;;;;;;;;;;", LIST_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				asColumnsTitles = Split("Folio;;;Area;;;F. del documento;;;F. límite;;;Ven;;;T. de Asunto;;;Prioridad;;;Estatus;;;Asunto", LIST_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				Do While Not oRecordset.EOF
                    iConta = iConta + 1
                    bVencido=False
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value) & ". Empleado: " & CStr(oRecordset.Fields("EmployeeID").Value))
                    sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
                    If CLng(oRecordset.Fields("EstimatedDate").Value) = 0 Then
					    sRowContents = sRowContents & TABLE_SEPARATOR & "ND"
                    Else
                        If (CLng(oRecordset.Fields("EstimatedDate").Value)<lStartDate) Then 
                            bVencido = True
                        End If
                        sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
                    End If
					If bVencido Then
                        sRowContents = sRowContents & TABLE_SEPARATOR & "X"
                    Else
                        sRowContents = sRowContents & TABLE_SEPARATOR & " "
                    End If
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
                    If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)
            lErrorNumber = AppendTextToFile(sFilePath, "<BR /><BR /><B>Número de registros: " & iConta & "</B><BR /><BR />", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1608 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1609(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1609"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber
    Dim sQuery
    Dim sSelectedDocs

	sQuery = CStr(oRequest("sQuery").Item)
    Response.Write sQuery & "<BR /><BR />"
    If Len(oRequest("SelectedDocs").Item) Then
        sSelectedDocs = "PaperworkID IN (" & CStr(oRequest("SelectedDocs").Item) & ")"
        sQuery = Replace(sQuery, "Order By PaperworkNumber", "And " & sSelectedDocs & " Order By PaperworkNumber")
    End If

    sCondition = ""
	'Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	'If Len(oRequest("YearID").Item) > 0 Then sCondition = sCondition & " And (Paperworks.StartDate>=" & oRequest("YearID").Item & "0000) And (Paperworks.StartDate<=" & oRequest("YearID").Item & "9999)"
	'If Len(oRequest("Closed").Item) > 0 Then
	'	If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then
	'		sCondition = sCondition & " And (Paperworks.EndDate=0)"
	'	Else
	'		sCondition = sCondition & " And (Paperworks.EndDate>0)"
	'	End If
	'End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		'If Not oRecordset.EOF Then
        If False Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<B>Fecha y  hora de la generación del reporte: " & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), Hour(Time()), Minute(Time()), Second(Time())) & "</B><BR /><BR />", sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("Folio;;;Procedencia;;;Fecha;;;Observaciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				Do While Not oRecordset.EOF
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & """)"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerEndDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "---"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("ClosingNumber").Value)) & """)"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "---"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value))
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "En trámite"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value))
						If CLng(oRecordset.Fields("AreaID").Value) > -1 Then sRowContents = sRowContents & CleanStringForHTML(" (" & CStr(oRecordset.Fields("AreaName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerAreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkActionName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			'Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
            Response.Write sQuery & "<BR /><BR />"
            Response.Write oRequest("SelectedDocs").Item & "<BR /><BR />"
            Response.Write "Update Paperworks Set List=1 Where " & sSelectedDocs
		End If
	End If
    'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Paperworks Set List=1 Where " & sSelectedDocs, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1609 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1609Full(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1609Full"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber
	Dim sQuery
	Dim iCount
	Dim sSelectedDocs
	Dim asSelectedDocs
	Dim asCurrentDoc
	Dim sDocID
	Dim sDocOwner
	Dim bFindDoc
	Dim sRequest

	sRequest = CStr(oRequest)
    'Response.Write sRequest & "<BR /><BR />"
	sQuery = CStr(oRequest("sQuery").Item)
    'Response.Write sQuery & "<BR /><BR />"
    If Len(oRequest("SelectedDocs").Item) > 0 Then
        asSelectedDocs = Split(CStr(oRequest("SelectedDocs").Item), ",")
    End If

    sCondition = ""
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
        'If False Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<B>Fecha y  hora de la generación del reporte: " & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), Hour(Time()), Minute(Time()), Second(Time())) & "</B><BR /><BR />", sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("Folio;;;Procedencia;;;Fecha;;;Responsable;;;Observaciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				Do While Not oRecordset.EOF
                    bFindDoc = False
                    For iCount = 0 To UBound(asSelectedDocs)
                        asCurrentDoc = Split(asSelectedDocs(iCount), LIST_SEPARATOR)
                        sDocID = Trim(asCurrentDoc(0))
                        sDocOwner = Trim(asCurrentDoc(1))
                        If (CStr(oRecordset.Fields("PaperworkID").Value)=sDocID) And (CStr(oRecordset.Fields("OwnerID").Value)=sDocOwner) Then
                            'Response.Write "Update PaperworkOwnersLKP Set List=1 Where PaperworkID=" & sDocID & " And OwnerID=" & sDocOwner & ";<BR />"
                            bFindDoc = True
                            Exit For
                        End If
                    Next
                    If bFindDoc Then
						sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value) & ". Empleado: " & CStr(oRecordset.Fields("EmployeeID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "" 'CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaperworkOwnersLKP Set List=1 Where PaperworkID=" & sDocID & " And OwnerID=" & sDocOwner, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
                    End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
            'Response.Write sQuery & "<BR /><BR />"
            'Response.Write oRequest("SelectedDocs").Item & "<BR /><BR />"
            'Response.Write "Update Paperworks Set List=1 Where " & sSelectedDocs
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1609Full = lErrorNumber
	Err.Clear
End Function

Function DisplayReport1609Table(oRequest, oADODBConnection, bForExport, sAction, sErrorDescription)
'*****************************************************************
'Purpose: To display the beneficiaries of employees
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReport1609Table"
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

    Dim sTables

	sTables = ""
	If InStr(1, sCondition, "Jobs", vbBinaryCompare) > 0 Then sTables = sTables & ", Jobs, Areas"
	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	bIsFirst = True
	sErrorDescription = "No existen beneficiarios de pención alimenticia para el empledado."
	sQuery = "Select Distinct Paperworks.*, StatusName, PriorityName, SubjectTypeName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks, Priorities, SubjectTypes" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) " & sCondition & " Order By PaperworkNumber"

	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DisplayReport1609Table.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
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
				asColumnsTitles = Split("Folio;;;F. del documento;;;F. límite;;;Asunto;;;Prioridad;;;Estatus;;;Observaciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("80,170,170,100,100,100,250", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Folio;;;F. del documento;;;F. límite;;;Asunto;;;Prioridad;;;Estatus;;;Observaciones;;;Acciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("80,170,170,100,100,100,250,100", ",", -1, vbBinaryCompare)
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
			asCellAlignments = Split(",,CENTER,,,,,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If bForExport Then
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
				Else
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
				End If
                sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
				If CLng(oRecordset.Fields("EstimatedDate").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("ND")
				Else
                    If (CLng(oRecordset.Fields("EstimatedDate").Value)<lStartDate) Then 
                        bVencido = True
                    End If
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
                sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
				If Not bForExport Then
					sRowContents = sRowContents & TABLE_SEPARATOR
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""SelectedDocs"" ID=""" & CStr(oRecordset.Fields("PaperworkID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("PaperworkID").Value) & """ CHECKED=""1"" &/>"
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
					sErrorDescription = "No existen registros de documentos que además no se hayan impreso que cumplan con los criterios del filtro"
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	DisplayReport1609Table = lErrorNumber
	Err.Clear
End Function

Function DisplayReport1609TableFull(oRequest, oADODBConnection, bForExport, sAction, sErrorDescription)
'*****************************************************************
'Purpose: To display the beneficiaries of employees
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReport1609TableFull"
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

    Dim sTables

	sTables = ""
	If InStr(1, sCondition, "Jobs", vbBinaryCompare) > 0 Then sTables = sTables & ", Jobs, Areas"
	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	bIsFirst = True
	sErrorDescription = "No existen beneficiarios de pención alimenticia para el empledado."
	sQuery = "Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName, PriorityName, SubjectTypeName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks, Priorities, SubjectTypes" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (PaperworkOwnersLKP.List=0)" & sCondition & " Order By PaperworkNumber"
    'sQuery = "Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName, PriorityName, SubjectTypeName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks, Priorities, SubjectTypes" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID)" & sCondition & " Order By PaperworkNumber"

	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DisplayReport1609Table.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			'If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine

			If bForExport Then
				asColumnsTitles = Split("Folio;;;Area;;;F. del documento;;;F. límite;;;Ven;;;Asunto;;;Prioridad;;;Estatus;;;Observaciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("80,200,170,170,100,100,100,100,250", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Folio;;;Area;;;F. del documento;;;F. límite;;;Ven;;;Asunto;;;Prioridad;;;Estatus;;;Observaciones;;;Acciones", LIST_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split("80,200,170,170,100,100,100,100,250,100", ",", -1, vbBinaryCompare)
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
			asCellAlignments = Split(",,CENTER,,,,,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If bForExport Then
					sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & """)"
				Else
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value) & ". Empleado: " & CStr(oRecordset.Fields("EmployeeID").Value))
                sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
				If CLng(oRecordset.Fields("EstimatedDate").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("ND")
				Else
                    If (CLng(oRecordset.Fields("EstimatedDate").Value)<lStartDate) Then 
                        bVencido = True
                    End If
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
				End If
                If bVencido Then
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("X")
                Else
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(" ")
                End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
                sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
				If Not bForExport Then
					sRowContents = sRowContents & TABLE_SEPARATOR
					If B_DELETE And (aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						'sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""SelectedDocs"" ID=""" & CStr(oRecordset.Fields("PaperworkNumber").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("PaperworkID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("OwnerID").Value) & """>"
						sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""SelectedDocs"" ID=""" & CStr(oRecordset.Fields("PaperworkNumber").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("PaperworkID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("OwnerID").Value) & """CHECKED=1 >"
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
					sErrorDescription = "No existen registros de documentos que además no se hayan impreso que cumplan con los criterios del filtro"
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	DisplayReport1609TableFull = lErrorNumber
	Err.Clear
End Function

Function BuildReport1607_old(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1607_old"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(oRequest("YearID").Item) > 0 Then sCondition = sCondition & " And (Paperworks.StartDate>=" & oRequest("YearID").Item & "0000) And (Paperworks.StartDate<=" & oRequest("YearID").Item & "9999)"
	If Len(oRequest("Closed").Item) > 0 Then
		If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then
			sCondition = sCondition & " And (Paperworks.EndDate=0)"
		Else
			sCondition = sCondition & " And (Paperworks.EndDate>0)"
		End If
	End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkTypeName, SubjectTypeName, StatusName, PriorityName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, Owners.EmployeeName As OwnerName1, Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName2 As OwnerLastName2, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, Areas.AreaID, Areas.AreaName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, Priorities, PaperworkOwnersLKP, PaperworkOwners, Employees As Owners, PaperworkActions, PaperworkSenders, Areas Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.EmployeeID=Owners.EmployeeID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkSenders.AreaID=Areas.AreaID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners.LevelID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PriorityName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, Owners.EmployeeName As OwnerName1, Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName2 As OwnerLastName2, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, Areas.AreaID, Areas.AreaName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, Priorities, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders, Areas Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PriorityID=Priorities.PriorityID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkSenders.AreaID=Areas.AreaID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners.LevelID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<B>Fecha y  hora de la generación del reporte:&nbsp;</B>" & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), Hour(Time()), Minute(Time()), Second(Time())) & "<BR /><BR />", sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split(",,,,,,", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = "C.C."
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha del documento"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha límite"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value)) & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Documento"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de recepción"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de descargo"
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerEndDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Prioridad"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PriorityName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha que se turnó"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "Oficio de descargo"
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ClosingNumber").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Remitente"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Estatus"
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Cerrado"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "En trámite"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "Documento de origen"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Procedencia"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""3"" />" & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value))
					If CLng(oRecordset.Fields("AreaID").Value) > -1 Then sRowContents = sRowContents & CleanStringForHTML(" (" & CStr(oRecordset.Fields("AreaName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR & "Responsable"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerAreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Tipo de asunto"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Asunto"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Observaciones"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Tipo de trámite"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)

					sRowContents = "Acciones"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkActionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1607_old = lErrorNumber
	Err.Clear
End Function

Function BuildReport1610(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1610"
	Dim iLevels
	Dim iCurrentOwnerID
	Dim lCounter
	Dim lTotal
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim bDone
	Dim lErrorNumber

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(oRequest("PaperworkLevelID").Item) > 0 Then sCondition = sCondition & " And (PaperworkOwners.LevelID<=" & oRequest("PaperworkLevelID").Item & ")"
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName As OwnerAreaName, PaperworkOwners.LevelID, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, PaperworkOwners2.OwnerID As OwnerID2, PaperworkOwners2.OwnerName As OwnerAreaName2, PaperworkOwners3.OwnerID As OwnerID3, PaperworkOwners3.OwnerName As OwnerAreaName3 From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders, PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners3 Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkOwners.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners3.OwnerID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners2.OwnerID, PaperworkOwners3.OwnerID, PaperworkOwners.OwnerID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName As OwnerAreaName, PaperworkOwners.LevelID, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName, PaperworkOwners2.OwnerID As OwnerID2, PaperworkOwners2.OwnerName As OwnerAreaName2, PaperworkOwners3.OwnerID As OwnerID3, PaperworkOwners3.OwnerName As OwnerAreaName3 From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders, PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners3 Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkOwners.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners3.OwnerID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber, PaperworkOwners2.OwnerID, PaperworkOwners3.OwnerID, PaperworkOwners.OwnerID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				Select Case oRequest("PaperworkLevelID").Item
					Case "1"
						asColumnsTitles = Split("Subdirección,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 0
					Case "2"
						asColumnsTitles = Split("Subdirección,Jefatura de servicio,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 1
					Case "3"
						asColumnsTitles = Split("Subdirección,Jefatura de servicio,Jefatura de depto.,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 2
				End Select
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				bDone = (CInt(oRecordset.Fields("LevelID").Value) > 1)
				lCounter = 0
				lTotal = 0
				iCurrentOwnerID = -2
				asCellAlignments = Split(",,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = ""
					Select Case CInt(oRecordset.Fields("LevelID").Value)
						Case 1
							If ((iCurrentOwnerID > -2) And (iLevels > 0)) Or bDone Then
								asRowContents = Split("<SPAN COLS=""" & 13 + iLevels & """ /><B>Asuntos: " & lCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
								lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
								lCounter = 0
								lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
							End If
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
							Select Case iLevels
								Case 1
									sRowContents = sRowContents & TABLE_SEPARATOR
								Case 2
									sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR
							End Select
							iCurrentOwnerID = CLng(oRecordset.Fields("OwnerID").Value)
							bDone = False
						Case 2
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID2").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName2").Value)) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
							Select Case iLevels
								Case 2
									sRowContents = sRowContents & TABLE_SEPARATOR
							End Select
						Case 3
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID3").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName3").Value)) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID2").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName2").Value)) & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))

					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value) & ". " & CStr(oRecordset.Fields("SenderName1").Value) & " (" & CStr(oRecordset.Fields("PositionName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Description").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkActionName").Value))
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerEndDate").Value), -1, -1, -1)
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CStr(oRecordset.Fields("OwnerEndDate").Value))) + 1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), Date()) + 1)
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					lCounter = lCounter + 1
					lTotal = lTotal + 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				asRowContents = Split("<SPAN COLS=""" & 13 + iLevels & """ /><B>Asuntos: " & lCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
				asRowContents = Split("<SPAN COLS=""" & 13 + iLevels & """ /><B>TOTAL: " & lTotal & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1610 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1611(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per month
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1611"
	Dim iMonth
	Dim iQuarter
	Dim lMonthCounter
	Dim lQuarterCounter
	Dim lYearCounter
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName From Paperworks, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkSenders Where (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.SenderID=PaperworkSenders.SenderID) " & sCondition & " Order By Paperworks.StartDate, PaperworkNumber", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkSenders Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.SenderID=PaperworkSenders.SenderID) " & sCondition & " Order By Paperworks.StartDate, PaperworkNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("Año,Trimestre,Mes,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de atención", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				asCellAlignments = Split(",,,,,,,,,,,,", ",", -1, vbBinaryCompare)

				iMonth = CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
				iQuarter = Int((iMonth - 1) / 3) + 1
				sRowContents = "<B>" & Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & iQuarter & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""11"" /><B>" & asMonthNames_es(iMonth) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lMonthCounter = 0
				lQuarterCounter = 0
				lYearCounter = 0

				Do While Not oRecordset.EOF
					If iMonth <> CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00"))) Then
						asRowContents = Split("<SPAN COLS=""13"" /><B>Asuntos recibidos en " & asMonthNames_es(iMonth) & ": " & lMonthCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
						iMonth = CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
						lMonthCounter = 0

						If iQuarter <> Int((iMonth - 1) / 3) + 1 Then
							asRowContents = Split("<SPAN COLS=""13"" /><B>Asuntos recibidos en el " & iQuarter & "o. trimestre : " & lQuarterCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
							iQuarter = Int((iMonth - 1) / 3) + 1
							lQuarterCounter = 0
						End If

						sRowContents = "<B>" & Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & iQuarter & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""11"" /><B>" & asMonthNames_es(iMonth) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					End If
					sRowContents = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value) & ". ")
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value))
					sRowContents = sRowContents & CleanStringForHTML(" (" & CStr(oRecordset.Fields("PositionName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Description").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					
					If CLng(oRecordset.Fields("EndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value))) + 1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), Date()) + 1)
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					lMonthCounter = lMonthCounter + 1
					lQuarterCounter = lQuarterCounter + 1
					lYearCounter = lYearCounter + 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				asRowContents = Split("<SPAN COLS=""13"" /><B>Asuntos recibidos en " & asMonthNames_es(iMonth) & ": " & lMonthCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				asRowContents = Split("<SPAN COLS=""13"" /><B>Asuntos recibidos en el " & iQuarter & "o. trimestre : " & lQuarterCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
				asRowContents = Split("<SPAN COLS=""13"" /><B>TOTAL: " & lYearCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1611 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1612(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks per owner and month
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1612"
	Dim iLevels
	Dim iMonth
	Dim iQuarter
	Dim lMonthCounter
	Dim lQuarterCounter
	Dim lYearCounter
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	sFilePath = Server.MapPath(sFileName & ".xls")

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(oRequest("PaperworkLevelID").Item) > 0 Then sCondition = sCondition & " And (PaperworkOwners.LevelID<=" & oRequest("PaperworkLevelID").Item & ")"
	If Len(oRequest("Closed").Item) > 0 Then
		If CInt(oRequest("Closed").Item) = 0 Then
			sCondition = sCondition & " And (PaperworkOwnersLKP.EndDate=0)"
			lErrorNumber = AppendTextToFile(sFilePath, "<B>ASUNTOS ABIERTOS</B>", sErrorDescription)
		Else
			sCondition = sCondition & " And (PaperworkOwnersLKP.EndDate>0)"
			lErrorNumber = AppendTextToFile(sFilePath, "<B>ASUNTOS CERRADOS</B>", sErrorDescription)
		End If
	Else
		lErrorNumber = AppendTextToFile(sFilePath, "<B>TODOS LOS ASUNTOS</B>", sErrorDescription)
	End If
	lErrorNumber = AppendTextToFile(sFilePath, "<BR /><BR />", sErrorDescription)

	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By Paperworks.StartDate, PaperworkNumber, PaperworkOwners.LevelID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Paperworks.*, PaperworkTypeName, SubjectTypeName, StatusName, PaperworkOwners.OwnerID, OwnerName As OwnerAreaName, PaperworkOwners.LevelID, PaperworkActionShortName, PaperworkActionName, ReportDate, PaperworkOwnersLKP.ReportDate As OwnerStartDate, PaperworkOwnersLKP.EndDate As OwnerEndDate, ClosingNumber, SenderName As SenderAreaName, PaperworkSenders.EmployeeName As SenderName1, PositionName From Paperworks, PaperworkTypes, SubjectTypes, StatusPaperworks, PaperworkOwnersLKP, PaperworkOwners, PaperworkActions, PaperworkSenders Where (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.SubjectTypeID=SubjectTypes.SubjectTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (Paperworks.SenderID=PaperworkSenders.SenderID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By Paperworks.StartDate, PaperworkNumber, PaperworkOwners.LevelID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				Select Case oRequest("PaperworkLevelID").Item
					Case "1"
						asColumnsTitles = Split("Año,Trimestre,Mes,Subdirección,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 0
					Case "2"
						asColumnsTitles = Split("Año,Trimestre,Mes,Subdirección,Jefatura de servicio,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 1
					Case "3"
						asColumnsTitles = Split("Año,Trimestre,Mes,Subdirección,Jefatura de servicio,Jefatura de depto.,No. de trámite,Fecha del documento,Documento,Procedencia,Desc. procedencia,Asunto,Tipo de trámite,Tipo de asunto,Observaciones,Fecha de turnado,Acción,Fecha de atención", ",", -1, vbBinaryCompare)
						iLevels = 2
				End Select
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)

				asCellAlignments = Split(",,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)

				iMonth = CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
				iQuarter = Int((iMonth - 1) / 3) + 1
				sRowContents = "<B>" & Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & iQuarter & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""" & 14 + iLevels & """ /><B>" & asMonthNames_es(iMonth) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lMonthCounter = 0
				lQuarterCounter = 0
				lYearCounter = 0

				Do While Not oRecordset.EOF
					If iMonth <> CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00"))) Then
						asRowContents = Split("<SPAN COLS=""" & 16 + iLevels & """ /><B>Asuntos recibidos en " & asMonthNames_es(iMonth) & ": " & lMonthCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
						iMonth = CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
						lMonthCounter = 0

						If iQuarter <> Int((iMonth - 1) / 3) + 1 Then
							asRowContents = Split("<SPAN COLS=""" & 16 + iLevels & """ /><B>Asuntos recibidos en el " & iQuarter & "o. trimestre : " & lQuarterCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
							iQuarter = Int((iMonth - 1) / 3) + 1
							lQuarterCounter = 0
						End If

						sRowContents = "<B>" & Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & iQuarter & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""" & 14 + iLevels & """ /><B>" & asMonthNames_es(iMonth) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					End If
					sRowContents = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
					Select Case CInt(oRecordset.Fields("LevelID").Value)
						Case 1
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
							Select Case iLevels
								Case 1
									sRowContents = sRowContents & TABLE_SEPARATOR
								Case 2
									sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR
							End Select
						Case 2
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
							Select Case iLevels
								Case 2
									sRowContents = sRowContents & TABLE_SEPARATOR
							End Select
						Case 3
							sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerAreaName").Value))
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderAreaName").Value) & ". ")
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value))
					sRowContents = sRowContents & CleanStringForHTML(" (" & CStr(oRecordset.Fields("PositionName").Value) & ")")
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Description").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerStartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkActionName").Value))
					If CLng(oRecordset.Fields("OwnerEndDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OwnerEndDate").Value), -1, -1, -1)
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CStr(oRecordset.Fields("OwnerEndDate").Value))) + 1)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
'						sRowContents = sRowContents & TABLE_SEPARATOR & (DateDiff("d", GetDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value)), Date()) + 1)
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					lMonthCounter = lMonthCounter + 1
					lQuarterCounter = lQuarterCounter + 1
					lYearCounter = lYearCounter + 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				asRowContents = Split("<SPAN COLS=""" & 16 + iLevels & """ /><B>Asuntos recibidos en " & asMonthNames_es(iMonth) & ": " & lMonthCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				asRowContents = Split("<SPAN COLS=""" & 16 + iLevels & """ /><B>Asuntos recibidos en el " & iQuarter & "o. trimestre : " & lQuarterCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "<TR><TD HEIGHT=""1""></TD></TR>", sErrorDescription)
				asRowContents = Split("<SPAN COLS=""" & 16 + iLevels & """ /><B>TOTAL: " & lYearCounter & "</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
			lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1612 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1613(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the number of paperworks by area
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1613"
	Dim lStartDate
	Dim lEndDate
	Dim iStep
	Dim asCondition
	Dim asCondition2
	Dim sCondition
	Dim sOwnerIDs
	Dim asTitles
	Dim oRecordset
	Dim asOwners
	Dim iIndex
	Dim jIndex
	Dim alTotals
	Dim alGlobalTotals
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	iStep = 1
	If Len(oRequest("Steps").Item) > 0 Then iStep = CInt(oRequest("Steps").Item)
	Call GetStartAndEndDatesFromURLAsNumbers("PaperworkStartStart", "PaperworkStartEnd", True, lStartDate, lEndDate)
	If lStartDate < CLng((Year(Date()) - 1) & "9999") Then
		asCondition = " And (Paperworks.StartDate<=" & Year(Date()) - 1 & "9999)" & LIST_SEPARATOR
		asCondition2 = "0.0.0." & Year(Date()) - 1 & ".99.99" & LIST_SEPARATOR
		asTitles = "PERIODOS ANTERIORES A " & Year(Date()) & LIST_SEPARATOR
	End If
	If lStartDate < CLng(Year(Date()) & "0101") Then
		iIndex = 1
	Else
		iIndex = CInt(Mid(lStartDate, Len("YYYYM"), Len("MM")))
	End If
	If lStartDate < CLng(Year(Date()) & "0101") Then lStartDate = CLng(Year(Date()) & "0101")
	For iIndex = CInt(Mid(lStartDate, Len("YYYYM"), Len("MM"))) To CInt(Mid(lEndDate, Len("YYYYM"), Len("MM"))) Step iStep
		asCondition = asCondition & "And (Paperworks.StartDate>=" & Year(Date()) & Right(("0" & iIndex), Len("00")) & "00) And (Paperworks.StartDate<=" & Year(Date()) & Right(("0" & (iIndex + (iStep - 1))), Len("00")) & "99)" & LIST_SEPARATOR
		asCondition2 = asCondition2 & Year(Date()) & "." & Right(("0" & iIndex), Len("00")) & ".00." & Year(Date()) & "." & Right(("0" & (iIndex + (iStep - 1))), Len("00")) & ".99" & LIST_SEPARATOR
		If iStep = 1 Then
			If iIndex = CInt(Mid(lStartDate, Len("YYYYM"), Len("MM"))) Then
				If CInt(Right(lStartDate, Len("DD"))) > 1 Then
					asTitles = asTitles & "PERIODO A PARTIR DEL " & Right(lStartDate, Len("DD")) & " DE " & UCase(asMonthNames_es(iIndex)) & " DE " & Year(Date()) & LIST_SEPARATOR
				Else
					asTitles = asTitles & "PERIODO " & UCase(asMonthNames_es(iIndex)) & " " & Year(Date()) & LIST_SEPARATOR
				End If
			ElseIf iIndex = CInt(Mid(lEndDate, Len("YYYYM"), Len("MM"))) Then
				asTitles = asTitles & "PERIODO HASTA EL " & Right(lEndDate, Len("DD")) & " DE " & UCase(asMonthNames_es(iIndex)) & " DE " & Year(Date()) & LIST_SEPARATOR
			Else
				asTitles = asTitles & "PERIODO " & UCase(asMonthNames_es(iIndex)) & " " & Year(Date()) & LIST_SEPARATOR
			End If
		Else
			asTitles = asTitles & "PERIODO " & UCase(asMonthNames_es(iIndex)) & " - " & UCase(asMonthNames_es(iIndex + (iStep-1))) & " " & Year(Date()) & LIST_SEPARATOR
		End If
	Next
	If Len(asCondition) > 0 Then asCondition = Left(asCondition, (Len(asCondition) - Len(LIST_SEPARATOR)))
	If Len(asCondition2) > 0 Then asCondition2 = Left(asCondition2, (Len(asCondition2) - Len(LIST_SEPARATOR)))
	If Len(asTitles) > 0 Then asTitles = Left(asTitles, (Len(asTitles) - Len(LIST_SEPARATOR)))
	asCondition = Split(asCondition, LIST_SEPARATOR)
	asCondition2 = Split(asCondition2, LIST_SEPARATOR)
	asTitles = Split(asTitles, LIST_SEPARATOR)

'	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(CStr(oRequest("OwnerIDs").Item)) = 0 Then
		Call GetPaperworksOwnersForUser(sOwnerIDs, "")
		If InStr(1, sOwnerIDs, "-1", vbBinaryCompare) = 0 Then sCondition = " And (PaperworkOwnersLKP.OwnerID In (" & sOwnerIDs & "))"
	Else
		sCondition = " And (PaperworkOwnersLKP.OwnerID In (" & oRequest("OwnerIDs").Item & "))"
	End If

	sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
	If InStr(1, sOwnerIDs, "-1", vbBinaryCompare) = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID, OwnerName From PaperworkOwners Where (OwnerID>-1) " & Replace(sCondition, "PaperworkOwnersLKP", "PaperworkOwners") & " Order By PaperworkOwners.OwnerID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwners2.OwnerID, PaperworkOwners2.OwnerName From PaperworkOwners As PaperworkOwners1, PaperworkOwners As PaperworkOwners2 Where (PaperworkOwners2.ParentID=PaperworkOwners1.OwnerID) And (PaperworkOwners1.ParentID=-1) And (PaperworkOwners2.ParentID>-1) And (PaperworkOwners2.OwnerID>-1) Order By PaperworkOwners1.OwnerID, PaperworkOwners2.OwnerID", "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			asOwners = ""
			alTotals = Split("0,0", ",")
			alGlobalTotals = Split("0,0", ",")

			Do While Not oRecordset.EOF
				asOwners = asOwners & CStr(oRecordset.Fields("OwnerID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("OwnerName").Value) & SECOND_LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR & "0" & LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			If Len(asOwners) > 0 Then asOwners = Left(asOwners, (Len(asOwners) - Len(LIST_SEPARATOR)))
			asOwners = Split(asOwners, LIST_SEPARATOR)

			alGlobalTotals(0) = 0
			alGlobalTotals(1) = 0
			For jIndex = 0 To UBound(asCondition)
				asCondition2(jIndex) = Split(asCondition2(jIndex), ".")
				alTotals(0) = 0
				alTotals(1) = 0
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>" & asTitles(jIndex) & "</B></FONT><BR />"
				Response.Write "<TABLE BORDER="""
					If bForExport Then
						Response.Write "1"
					Else
						Response.Write "0"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"

					asColumnsTitles = Split("Área,Asuntos registrados,Asuntos descargados,Pendiendes de descargo,Porcentaje de pendientes", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100", ",", -1, vbBinaryCompare)
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
					For iIndex = 0 To UBound(asOwners)
						asOwners(iIndex) = Split(asOwners(iIndex), SECOND_LIST_SEPARATOR)
						asOwners(iIndex)(2) = 0
						asOwners(iIndex)(3) = 0

						sRowContents = CleanStringForHTML(asOwners(iIndex)(1))
						If Not bForExport Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""EmployeeSupport.asp?StartStartDay=" & asCondition2(jIndex)(2) & "&StartStartMonth=" & asCondition2(jIndex)(1) & "&StartStartYear=" & asCondition2(jIndex)(0) & "&EndStartDay=" & asCondition2(jIndex)(5) & "&EndStartMonth=" & asCondition2(jIndex)(4) & "&EndStartYear=" & asCondition2(jIndex)(3) & "&FilterOwnerID=" & asOwners(iIndex)(0) & "&FullSearch=1&DoSearch=1""><TOTAL /></A>" & TABLE_SEPARATOR
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<TOTAL />" & TABLE_SEPARATOR
						End If
						sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As TotalPaperworks From Paperworks, PaperworkOwnersLKP, PaperworkOwners As PaperworkOwners1, PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners3 Where (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners3.OwnerID) And (PaperworkOwners3.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners1.OwnerID) And (PaperworkOwners3.OwnerID>-1) And (PaperworkOwnersLKP.EndDate<>0) And ((PaperworkOwners1.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners2.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners3.OwnerID=" & asOwners(iIndex)(0) & ")) " & asCondition(jIndex) & sCondition, "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write vbNewLine & "<!-- Query: Select Count(*) As TotalPaperworks From Paperworks, PaperworkOwnersLKP, PaperworkOwners As PaperworkOwners1, PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners3 Where (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners3.OwnerID) And (PaperworkOwners3.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners1.OwnerID) And (PaperworkOwners3.OwnerID>-1) And (PaperworkOwnersLKP.EndDate<>0) And ((PaperworkOwners1.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners2.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners3.OwnerID=" & asOwners(iIndex)(0) & ")) " & asCondition(jIndex) & "-->" & vbNewLine
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								If Not IsNull(oRecordset.Fields("TotalPaperworks").Value) Then
									If Not bForExport Then sRowContents = sRowContents & "<A HREF=""EmployeeSupport.asp?StartStartDay=" & asCondition2(jIndex)(2) & "&StartStartMonth=" & asCondition2(jIndex)(1) & "&StartStartYear=" & asCondition2(jIndex)(0) & "&EndStartDay=" & asCondition2(jIndex)(5) & "&EndStartMonth=" & asCondition2(jIndex)(4) & "&EndStartYear=" & asCondition2(jIndex)(3) & "&FilterOwnerID=" & asOwners(iIndex)(0) & "&Closed=1&FullSearch=1&DoSearch=1"">"
										sRowContents = sRowContents & FormatNumber(CLng(oRecordset.Fields("TotalPaperworks").Value), 0, True, False, True)
									If Not bForExport Then sRowContents = sRowContents & "</A>"
									asOwners(iIndex)(2) = CLng(oRecordset.Fields("TotalPaperworks").Value)
									alTotals(0) = alTotals(0) + CLng(oRecordset.Fields("TotalPaperworks").Value)
									alGlobalTotals(0) = alGlobalTotals(0) + CLng(oRecordset.Fields("TotalPaperworks").Value)
								Else
									sRowContents = sRowContents & "0"
								End If
							Else
								sRowContents = sRowContents & "0"
							End If
						Else
							sRowContents = sRowContents & "0"
						End If

						sRowContents = sRowContents & TABLE_SEPARATOR
						sErrorDescription = "No se pudieron obtener los trámites registrados en el sistema."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As TotalPaperworks From Paperworks, PaperworkOwnersLKP, PaperworkOwners As PaperworkOwners1, PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners3 Where (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners3.OwnerID) And (PaperworkOwners3.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners1.OwnerID) And (PaperworkOwners3.OwnerID>-1) And (PaperworkOwnersLKP.EndDate=0) And ((PaperworkOwners1.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners2.OwnerID=" & asOwners(iIndex)(0) & ") Or (PaperworkOwners3.OwnerID=" & asOwners(iIndex)(0) & ")) " & asCondition(jIndex), "ReportsQueries1500Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								If Not IsNull(oRecordset.Fields("TotalPaperworks").Value) Then
									If Not bForExport Then sRowContents = sRowContents & "<A HREF=""EmployeeSupport.asp?StartStartDay=" & asCondition2(jIndex)(2) & "&StartStartMonth=" & asCondition2(jIndex)(1) & "&StartStartYear=" & asCondition2(jIndex)(0) & "&EndStartDay=" & asCondition2(jIndex)(5) & "&EndStartMonth=" & asCondition2(jIndex)(4) & "&EndStartYear=" & asCondition2(jIndex)(3) & "&FilterOwnerID=" & asOwners(iIndex)(0) & "&Closed=0&FullSearch=1&DoSearch=1"">"
										sRowContents = sRowContents & FormatNumber(CLng(oRecordset.Fields("TotalPaperworks").Value), 0, True, False, True)
									If Not bForExport Then sRowContents = sRowContents & "</A>"
									asOwners(iIndex)(3) = CLng(oRecordset.Fields("TotalPaperworks").Value)
									alTotals(1) = alTotals(1) + CLng(oRecordset.Fields("TotalPaperworks").Value)
									alGlobalTotals(1) = alGlobalTotals(1) + CLng(oRecordset.Fields("TotalPaperworks").Value)
								Else
									sRowContents = sRowContents & "0"
								End If
							Else
								sRowContents = sRowContents & "0"
							End If
						Else
							sRowContents = sRowContents & "0"
						End If

						sRowContents = sRowContents & TABLE_SEPARATOR
						If asOwners(iIndex)(2) = 0 Then
							sRowContents = sRowContents & "0.00%"
						Else
							sRowContents = sRowContents & FormatNumber(((asOwners(iIndex)(3) / (asOwners(iIndex)(2) + asOwners(iIndex)(3))) * 100), 2, True, False, True) & "%"
						End If
						sRowContents = Replace(sRowContents, "<TOTAL />", FormatNumber((asOwners(iIndex)(2) + asOwners(iIndex)(3)), 0, True, False, True))

						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					Next
					sRowContents = "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber((alTotals(0) + alTotals(1)), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(alTotals(0), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(alTotals(1), 0, True, False, True) & "</B>"
					sRowContents = sRowContents & TABLE_SEPARATOR
					If alTotals(1) = 0 Then
						sRowContents = sRowContents & "<B>0.00%</B>"
					Else
						sRowContents = sRowContents & "<B>" & FormatNumber(((alTotals(1) / (alTotals(0) + alTotals(1))) * 100), 2, True, False, True) & "%</B>"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				Response.Write "</TABLE><BR /><BR />"
			Next

			Response.Write "<B>TOTAL DE ASUNTOS REGISTRADOS: " & FormatNumber((alGlobalTotals(0) + alGlobalTotals(1)), 0, True, False, True) & "</B><BR />"
			Response.Write "<B>TOTAL DE ASUNTOS DESCARGADOS: " & FormatNumber(alGlobalTotals(0), 0, True, False, True) & "</B><BR />"
			Response.Write "<B>TOTAL DE ASUNTOS PENDIENTES DE DESCARGO: " & FormatNumber(alGlobalTotals(1), 0, True, False, True) & "</B><BR />"
		Else
			Response.Write "&nbsp;&nbsp;&nbsp;<B>No existen registros en la base de datos que cumplan con los criterios del filtro.</B>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1613 = lErrorNumber
	Err.Clear
End Function
%>