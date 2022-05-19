<%
Function BuildReport1100(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the report specially requested by ISSSTE
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1100"
	Dim sCondition
	Dim sTableNames
	Dim sJoinCondition
	Dim sDate
	Dim oRecordset
	Dim asColumnsTitles
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFileName
	Dim sFilePath
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	Call GetDBFieldsNames(oRequest, -1, sCondition, "", sTableNames, sJoinCondition, "")
	If Len(Trim(sTableNames)) > 0 Then sTableNames = ", " & sTableNames

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID From Employees " & sTableNames & " Where (EmployeeID>-1) " & sCondition & sJoinCondition & " Order By EmployeeNumber", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
			sFilePath = Server.MapPath(sFileName)
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFileName, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			Do While Not oRecordset.EOF
				sRowContents = SizeText("01", " ", 2, 1) 'Tipo de registro
				sRowContents = sRowContents & SizeText("09", " ", 2, 1) 'Identificador del Servicio
				sRowContents = sRowContents & SizeText("96", " ", 2, 1) 'Identificador de la Operación
				sRowContents = sRowContents & SizeText("XX", " ", 2, 1) 'Tipo entidad origen
				sRowContents = sRowContents & SizeText("XXXXXXX", " ", 7, 1) 'Clave entidad origen
				sRowContents = sRowContents & SizeText("XX", " ", 2, 1) 'Tipo entidad destino
				sRowContents = sRowContents & SizeText("XXXXXXX", " ", 7, 1) 'Clave entidad destino
				sRowContents = sRowContents & SizeText("XXXXXXXX", " ", 8, 1) 'Fecha de Transmisión

				sRowContents = sRowContents & SizeText("XXXXXXXXXXXX", " ", 12, 1) 'RFC de la Dependencia o Entidad con Homoclave
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 130, 1) 'Nombre de la Dependencia, Entidad o Cantro de Pago
				sRowContents = sRowContents & SizeText("XX", " ", 7, 1) 'Identificador de Centro de Pago SAR
				sRowContents = sRowContents & SizeText("XX", " ", 5, 1) 'Clave del Ramo
				sRowContents = sRowContents & SizeText("XX", " ", 5, 1) 'Clave de la Pagaduría
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 40, 1) 'Domicilio (Calle y número)
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 25, 1) 'Colonia
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 25, 1) 'Población, Delegación o Municipio
				sRowContents = sRowContents & SizeText("XXXXX", " ", 5, 1) 'Código Postal
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 23, 1) 'Entidad Federativa
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 10, 1) 'Teléfono

				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 9, 1) 'Total de registros con movimientos de Alta
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 9, 1) 'Total de registros con movimientos de Modificaciones
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 9, 1) 'Total de registros con movimientos de Bajas
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 9, 1) 'Total de registros de Detalle

				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 264, 1) 'Filler
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 2, 1) 'Resultado de la Operación
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 3, 1) 'Motivo de Rechazo 1
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 3, 1) 'Motivo de Rechazo 2
				sRowContents = sRowContents & SizeText("XXXXXXXXXX", " ", 3, 1) 'Motivo de Rechazo 3

				lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudo guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFileName, ".txt", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1100 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1101(oRequest, oADODBConnection, lYearID, iBonusType, iForReport, sErrorDescription)
'************************************************************
'Purpose: Reporte de validación del pago de aguinaldo
'Inputs:  oRequest, oADODBConnection, lYearID, iBonusType, iForReport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1101"
	Dim sCondition
	Dim oRecordset
	Dim lCurrentID
	Dim sCurrentID
	Dim sTempCurrentID
	Dim lStartDate
	Dim lEndDate
	Dim lTempStartDate
	Dim lTempEndDate
	Dim lEmployeeStartDate
	Dim lEmployeeEndDate
	Dim adAmounts
	Dim adAmounts3
	Dim adAmounts11
	Dim dAmounts
	Dim dAmounts3
	Dim dAmounts11
	Dim aiDays
	Dim aDays
	Dim iDays
	Dim iIndex
	Dim sTemp
	Dim bContinuous
	Dim bContinuous3
	Dim bContinuous11
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim asColumnsTitles
	Dim sRowContents
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	
	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(Replace(Replace(sCondition, "Areas.", "Areas2."), "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "Employees.", "EmployeesHistoryList.")
	lStartDate = CLng(lYearID & "0101")
	lEndDate = CLng(lYearID & "1231")
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If

	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JobID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.AreaID, EmployeesHistoryList.WorkingHours, Areas2.EconomicZoneID, EmployeesHistoryList.PositionID, EmployeesHistoryList.LevelID From EmployeesHistoryList, StatusEmployees, Reasons, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=EmployeesHistoryList.EmployeeDate) And (Areas1.EndDate>=EmployeesHistoryList.EmployeeDate) And (Areas2.StartDate<=EmployeesHistoryList.EmployeeDate) And (Areas2.EndDate>=EmployeesHistoryList.EmployeeDate) And (PaymentCenters.StartDate<=EmployeesHistoryList.EmployeeDate) And (PaymentCenters.EndDate>=EmployeesHistoryList.EmployeeDate) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & "))) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber, EmployeesHistoryList.EmployeeDate", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JobID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.AreaID, EmployeesHistoryList.WorkingHours, Areas2.EconomicZoneID, EmployeesHistoryList.PositionID, EmployeesHistoryList.LevelID From EmployeesHistoryList, StatusEmployees, Reasons, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (Areas1.StartDate<=EmployeesHistoryList.EmployeeDate) And (Areas1.EndDate>=EmployeesHistoryList.EmployeeDate) And (Areas2.StartDate<=EmployeesHistoryList.EmployeeDate) And (Areas2.EndDate>=EmployeesHistoryList.EmployeeDate) And (PaymentCenters.StartDate<=EmployeesHistoryList.EmployeeDate) And (PaymentCenters.EndDate>=EmployeesHistoryList.EmployeeDate) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & "))) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber, EmployeesHistoryList.EmployeeDate -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If iForReport = 1 Then
				sDate = GetSerialNumberForDate("")
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
			End If

			lCurrentID = -2
			sCurrentID = ""
			aiDays = Split("0,0,0,0", ",")
			For iIndex = 0 To UBound(aiDays)
				aiDays(iIndex) = 0
			Next
			Do While Not oRecordset.EOF
				sTempCurrentID = CStr(oRecordset.Fields("EmployeeID").Value) & ";;;" & CStr(oRecordset.Fields("JobID").Value) & ";;;" & CStr(oRecordset.Fields("PositionID").Value) & ";;;" & CStr(oRecordset.Fields("LevelID").Value) & ";;;" & CStr(oRecordset.Fields("WorkingHours").Value) & ";;;" & CStr(oRecordset.Fields("EconomicZoneID").Value)
				If StrComp(sCurrentID, sTempCurrentID, vbBinaryCompare) <> 0 Then
					If Len(sCurrentID) > 0 Then
						sRowContents = Replace(sRowContents, "<START_DATE />", lEmployeeStartDate)
						sRowContents = Replace(sRowContents, "<END_DATE />", lEmployeeEndDate)
						aDays = aiDays(0) / 30.4
						aDays = Split(aDays, ".")
						aDays(0) = CInt(aDays(0))
						aDays(1) = CDbl("0." & aDays(1))
						sRowContents = Replace(sRowContents, "<DAYS />", ((aDays(0) * 30) + CInt(aDays(1) * 30.4)))
						lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), sRowContents, sErrorDescription)
					End If

					sRowContents = CStr(oRecordset.Fields("EmployeeID").Value)
					sRowContents = sRowContents & "," & iForReport
					sRowContents = sRowContents & ",<START_DATE />,<END_DATE />"
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("JobID").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("PositionID").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("AreaID").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EconomicZoneID").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("LevelID").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("WorkingHours").Value)
					sRowContents = sRowContents & ",<DAYS />,0,0,0,0,0"

					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					sCurrentID = sTempCurrentID
					lEmployeeStartDate = 30000000
					lEmployeeEndDate = 0
					aiDays = Split("0,0,0,0", ",")
					For iIndex = 0 To UBound(aiDays)
						aiDays(iIndex) = 0
					Next
				End If
				If lStartDate > CLng(oRecordset.Fields("EmployeeDate").Value) Then
					lTempStartDate = lStartDate
				Else
					lTempStartDate = CLng(oRecordset.Fields("EmployeeDate").Value)
				End If
				If lEndDate < CLng(oRecordset.Fields("EndDate").Value) Then
					lTempEndDate = lEndDate
				Else
					lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
				End If
				If lEmployeeStartDate > lTempStartDate Then lEmployeeStartDate = lTempStartDate
				If lEmployeeEndDate < lTempEndDate Then lEmployeeEndDate = lTempEndDate
				aiDays(0) = aiDays(0) + Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			sRowContents = Replace(sRowContents, "<START_DATE />", lEmployeeStartDate)
			sRowContents = Replace(sRowContents, "<END_DATE />", lEmployeeEndDate)
			aDays = aiDays(0) / 30.4
			aDays = Split(aDays, ".")
			aDays(0) = CInt(aDays(0))
			aDays(1) = CDbl("0." & aDays(1))
			sRowContents = Replace(sRowContents, "<DAYS />", ((aDays(0) * 30) + CInt(aDays(1) * 30.4)))
			lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), sRowContents, sErrorDescription)

			sErrorDescription = "No se pudo borrar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesXmasBonus Where (bForReport=" & iForReport & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				asRowContents = GetFileContents(Server.MapPath(sFileName & ".txt"), sErrorDescription)
				asRowContents = Split(asRowContents, vbNewLine)
				For iIndex = 0 To UBound(asRowContents)
					If Len(asRowContents(iIndex)) > 0 Then
						sErrorDescription = "No se pudo guardar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesXmasBonus (EmployeeID, bForReport, StartDate, EndDate, JobID, PositionID, AreaID, EconomicZoneID, LevelID, WorkingHours, Days1, Days2, Days3, TotalAmount, TotalAmount3, TotalAmount11) Values (" & asRowContents(iIndex) & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				Next
			End If
			Call DeleteFile(Server.MapPath(sFileName & ".txt"), "")

			sErrorDescription = "No se pudieron obtener los registros de la base de datos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesXmasBonus.EmployeeID, EmployeesXmasBonus.StartDate As lStartDate, EmployeesXmasBonus.EndDate As lEndDate, EmployeesAbsencesLKP.AbsenceID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, EmployeesXmasBonus Where (EmployeesAbsencesLKP.EmployeeID=EmployeesXmasBonus.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID In (10)) And (((EmployeesAbsencesLKP.EndDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.EndDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.OcurredDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.EndDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate<=EmployeesXmasBonus.EndDate) And (EmployeesAbsencesLKP.EndDate>=EmployeesXmasBonus.StartDate))) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) Order By EmployeesXmasBonus.EmployeeID, EmployeesXmasBonus.StartDate, EmployeesAbsencesLKP.OcurredDate", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesXmasBonus.EmployeeID, EmployeesXmasBonus.StartDate As lStartDate, EmployeesXmasBonus.EndDate As lEndDate, EmployeesAbsencesLKP.AbsenceID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, EmployeesXmasBonus Where (EmployeesAbsencesLKP.EmployeeID=EmployeesXmasBonus.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID In (10)) And (((EmployeesAbsencesLKP.EndDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.EndDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.OcurredDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate>=EmployeesXmasBonus.StartDate) And (EmployeesAbsencesLKP.EndDate<=EmployeesXmasBonus.EndDate)) Or ((EmployeesAbsencesLKP.OcurredDate<=EmployeesXmasBonus.EndDate) And (EmployeesAbsencesLKP.EndDate>=EmployeesXmasBonus.StartDate))) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) Order By EmployeesXmasBonus.EmployeeID, EmployeesXmasBonus.StartDate, EmployeesAbsencesLKP.OcurredDate -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lCurrentID = -2
					sCurrentID = ""
					For iIndex = 0 To UBound(aiDays)
						aiDays(iIndex) = 0
					Next
					Do While Not oRecordset.EOF
						sTempCurrentID = "(EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (StartDate=" & CStr(oRecordset.Fields("lStartDate").Value) & ")"
						If StrComp(sCurrentID, sTempCurrentID, vbBinaryCompare) <> 0 Then
							If Len(sCurrentID) > 0 Then
								lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), "Update EmployeesXmasBonus Set Days2=" & aiDays(1) & ", Days3=" & aiDays(2) & " Where " & sCurrentID, sErrorDescription)
							End If
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							sCurrentID = sTempCurrentID
							For iIndex = 0 To UBound(aiDays)
								aiDays(iIndex) = 0
							Next
						End If
						If CLng(oRecordset.Fields("lStartDate").Value) > CLng(oRecordset.Fields("OcurredDate").Value) Then
							lTempStartDate = CLng(oRecordset.Fields("lStartDate").Value)
						Else
							lTempStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
						End If
						If CLng(oRecordset.Fields("lEndDate").Value) < CLng(oRecordset.Fields("EndDate").Value) Then
							lTempEndDate = CLng(oRecordset.Fields("lEndDate").Value)
						Else
							lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
						End If
						Select Case CLng(oRecordset.Fields("AbsenceID").Value)
							Case 10
								aiDays(2) = aiDays(2) + Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1
							Case Else
								aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1
						End Select
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
					lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), "Update EmployeesXmasBonus Set Days2=" & aiDays(1) & ", Days3=" & aiDays(2) & " Where " & sCurrentID, sErrorDescription)

					asRowContents = GetFileContents(Server.MapPath(sFileName & ".txt"), sErrorDescription)
					asRowContents = Split(asRowContents, vbNewLine)
					For iIndex = 0 To UBound(asRowContents)
						If Len(asRowContents(iIndex)) > 0 Then
							sErrorDescription = "No se pudo actualizar la información del registro."
							lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, asRowContents(iIndex), "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
						End If
					Next
					Call DeleteFile(Server.MapPath(sFileName & ".txt"), "")
				End If
			End If

			sErrorDescription = "No se pudo borrar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select Max(Payroll_" & lYearID & ".RecordID) As RecordDate, EmployeesXmasBonus.StartDate As RecordID, EmployeesXmasBonus.EmployeeID, 0 As ConceptID, 0 As PayrollTypeID, 0 As ConceptAmount, 0 As ConceptTaxes, 0 As ConceptRetention, 0 As UserID From EmployeesXmasBonus, Payroll_" & lYearID & " Where (EmployeesXmasBonus.EmployeeID=Payroll_" & lYearID & ".EmployeeID) And (EmployeesXmasBonus.StartDate<=Payroll_" & lYearID & ".RecordID) And (EmployeesXmasBonus.EndDate>=Payroll_" & lYearID & ".RecordID) Group By EmployeesXmasBonus.EmployeeID, EmployeesXmasBonus.StartDate", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If (iBonusType = -1) Or (iBonusType = 1) Then
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payroll.EmployeeID, Payroll.RecordID, Sum(Payroll_" & lYearID & ".ConceptAmount) As TotalAmount From Payroll, Payroll_" & lYearID & " Where (Payroll.EmployeeID=Payroll_" & lYearID & ".EmployeeID) And (Payroll.RecordDate=Payroll_" & lYearID & ".RecordID) And (Payroll_" & lYearID & ".ConceptID In (1,7,8)) Group by Payroll.EmployeeID, Payroll.RecordID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Do While Not oRecordset.EOF
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), "Update EmployeesXmasBonus Set TotalAmount=" & CStr(oRecordset.Fields("TotalAmount").Value) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (bForReport=" & iForReport & ") And (StartDate=" & CStr(oRecordset.Fields("RecordID").Value) & ")", sErrorDescription)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
								asRowContents = GetFileContents(Server.MapPath(sFileName & ".txt"), sErrorDescription)
								asRowContents = Split(asRowContents, vbNewLine)
								For iIndex = 0 To UBound(asRowContents)
									If Len(asRowContents(iIndex)) > 0 Then
										sErrorDescription = "No se pudo actualizar la información del registro."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, asRowContents(iIndex), "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								Next
								Call DeleteFile(Server.MapPath(sFileName & ".txt"), "")
							Else
								oRecordset.Close
							End If
						End If
					End If
				End If

				If (iBonusType = -1) Or (iBonusType = 3) Then
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payroll.EmployeeID, Payroll.RecordID, Sum(Payroll_" & lYearID & ".ConceptAmount) As TotalAmount From Payroll, Payroll_" & lYearID & " Where (Payroll.EmployeeID=Payroll_" & lYearID & ".EmployeeID) And (Payroll.RecordDate=Payroll_" & lYearID & ".RecordID) And (Payroll_" & lYearID & ".ConceptID In (3)) Group by Payroll.EmployeeID, Payroll.RecordID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Do While Not oRecordset.EOF
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), "Update EmployeesXmasBonus Set TotalAmount3=" & CStr(oRecordset.Fields("TotalAmount").Value) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (bForReport=" & iForReport & ") And (StartDate=" & CStr(oRecordset.Fields("RecordID").Value) & ")", sErrorDescription)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
								asRowContents = GetFileContents(Server.MapPath(sFileName & ".txt"), sErrorDescription)
								asRowContents = Split(asRowContents, vbNewLine)
								For iIndex = 0 To UBound(asRowContents)
									If Len(asRowContents(iIndex)) > 0 Then
										sErrorDescription = "No se pudo actualizar la información del registro."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, asRowContents(iIndex), "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								Next
								Call DeleteFile(Server.MapPath(sFileName & ".txt"), "")
							Else
								oRecordset.Close
							End If
						End If
					End If
				End If

				If (iBonusType = -1) Or (iBonusType = 11) Then
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payroll.EmployeeID, Payroll.RecordID, Sum(Payroll_" & lYearID & ".ConceptAmount) As TotalAmount From Payroll, Payroll_" & lYearID & " Where (Payroll.EmployeeID=Payroll_" & lYearID & ".EmployeeID) And (Payroll.RecordDate=Payroll_" & lYearID & ".RecordID) And (Payroll_" & lYearID & ".ConceptID In (13)) Group by Payroll.EmployeeID, Payroll.RecordID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Do While Not oRecordset.EOF
									lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), "Update EmployeesXmasBonus Set TotalAmount11=" & CStr(oRecordset.Fields("TotalAmount").Value) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (bForReport=" & iForReport & ") And (StartDate=" & CStr(oRecordset.Fields("RecordID").Value) & ")", sErrorDescription)
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
								oRecordset.Close
								asRowContents = GetFileContents(Server.MapPath(sFileName & ".txt"), sErrorDescription)
								asRowContents = Split(asRowContents, vbNewLine)
								For iIndex = 0 To UBound(asRowContents)
									If Len(asRowContents(iIndex)) > 0 Then
										sErrorDescription = "No se pudo actualizar la información del registro."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, asRowContents(iIndex), "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								Next
								Call DeleteFile(Server.MapPath(sFileName & ".txt"), "")
							Else
								oRecordset.Close
							End If
						End If
					End If
				End If
			End If
			sErrorDescription = "No se pudo borrar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			If iForReport = 1 Then
				lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".xls"), "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
					asColumnsTitles = Split("No. Empleado;Empleado;RFC;Fecha de inicio;Fecha final;Estatus;Plaza;Puesto;Adscripción;Zona económica;Nivel;Jornada;Días laborados;Faltas;Días de aguinaldo;Días para pago;Pago completo (01, 07, 08);Pago proporcional (01, 07, 08);Aguinaldo 03;Aguinaldo 11", ";", -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".xls"), GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)
					asCellAlignments = Split(",,,,,,,,,CENTER,,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
					sErrorDescription = "No se pudieron obtener los registros de la base de datos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesXmasBonus.*, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, StatusName, AreaCode, PositionShortName From Employees, EmployeesXmasBonus, StatusEmployees, Positions, Areas Where (Employees.EmployeeID=EmployeesXmasBonus.EmployeeID) And (Employees.StatusID=StatusEmployees.StatusID) And (EmployeesXmasBonus.PositionID=Positions.PositionID) And (EmployeesXmasBonus.AreaID=Areas.AreaID) And (Areas.StartDate<=EmployeesXmasBonus.StartDate) And (Areas.EndDate>=EmployeesXmasBonus.StartDate) And (Positions.StartDate<=EmployeesXmasBonus.StartDate) And (Positions.EndDate>=EmployeesXmasBonus.StartDate) And (EmployeesXmasBonus.bForReport=" & iForReport & ") Order By EmployeeNumber, EmployeesXmasBonus.StartDate", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select EmployeesXmasBonus.*, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, StatusName, AreaCode, PositionShortName From Employees, EmployeesXmasBonus, StatusEmployees, Positions, Areas Where (Employees.EmployeeID=EmployeesXmasBonus.EmployeeID) And (Employees.StatusID=StatusEmployees.StatusID) And (EmployeesXmasBonus.PositionID=Positions.PositionID) And (EmployeesXmasBonus.AreaID=Areas.AreaID) And (Areas.StartDate<=EmployeesXmasBonus.StartDate) And (Areas.EndDate>=EmployeesXmasBonus.StartDate) And (Positions.StartDate<=EmployeesXmasBonus.StartDate) And (Positions.EndDate>=EmployeesXmasBonus.StartDate) And (EmployeesXmasBonus.bForReport=" & iForReport & ") Order By EmployeeNumber, EmployeesXmasBonus.StartDate -->" & vbNewLine
					If lErrorNumber = 0 Then
						lCurrentID = -2
						aiDays = ""
						adAmounts = ""
						adAmounts3 = ""
						adAmounts11 = ""
						lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
						dAmounts = 0
						dAmounts3 = 0
						dAmounts11 = 0
						Do While Not oRecordset.EOF
							lTempStartDate = AddDaysToSerialDate(CLng(oRecordset.Fields("StartDate").Value), -1)
							If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								If lCurrentID > -2 Then
									aiDays = Split(aiDays, LIST_SEPARATOR)
									adAmounts = Split(adAmounts, LIST_SEPARATOR)
									adAmounts3 = Split(adAmounts3, LIST_SEPARATOR)
									adAmounts11 = Split(adAmounts11, LIST_SEPARATOR)
									sRowContents = GetFileContents(Server.MapPath(sFileName & ".xls"), sErrorDescription)
									If bContinuous Then
										iDays = 0
										dAmounts = 0
										For iIndex = 0 To UBound(aiDays) - 2
											sRowContents = Replace(sRowContents, "<DAYS />", "0", 1, 1, vbBinaryCompare)
											sRowContents = Replace(sRowContents, "<AMOUNT />", "0.00", 1, 1, vbBinaryCompare)
											sRowContents = Replace(sRowContents, "<AMOUNT />", "0.00", 1, 1, vbBinaryCompare)
											iDays = iDays + CInt(aiDays(iIndex))
											dAmounts = dAmounts + CDbl(adAmounts(iIndex))
										Next
										iDays = iDays + CInt(aiDays(iIndex))
										dAmounts = dAmounts + CDbl(adAmounts(iIndex))
										If iDays > 360 Then iDays = 360
										sRowContents = Replace(sRowContents, "<DAYS />", iDays, 1, 1, vbBinaryCompare)
										sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((dAmounts / 15 * 40), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
										sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((dAmounts / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
									Else
										For iIndex = 0 To UBound(aiDays) - 1
											sRowContents = Replace(sRowContents, "<DAYS />", aiDays(iIndex), 1, 1, vbBinaryCompare)
											sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((adAmounts(iIndex) / 15 * 40), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
											sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((adAmounts(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
										Next
									End If
									If bContinuous3 Then
										iDays = 0
										dAmounts3 = 0
										For iIndex = 0 To UBound(aiDays) - 2
											sRowContents = Replace(sRowContents, "<AMOUNT_3 />", "0.00", 1, 1, vbBinaryCompare)
											iDays = iDays + CInt(aiDays(iIndex))
											dAmounts3 = dAmounts3 + CDbl(adAmounts3(iIndex))
										Next
										iDays = iDays + CInt(aiDays(iIndex))
										dAmounts3 = dAmounts3 + CDbl(adAmounts3(iIndex))
										If iDays > 360 Then iDays = 360
										sRowContents = Replace(sRowContents, "<AMOUNT_3 />", FormatNumber(FormatNumber((dAmounts3 / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
									Else
										For iIndex = 0 To UBound(aiDays) - 1
											sRowContents = Replace(sRowContents, "<AMOUNT_3 />", FormatNumber(FormatNumber((adAmounts3(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
										Next
									End If
									If bContinuous11 Then
										iDays = 0
										dAmounts11 = 0
										For iIndex = 0 To UBound(aiDays) - 2
											sRowContents = Replace(sRowContents, "<AMOUNT_11 />", "0.00", 1, 1, vbBinaryCompare)
											iDays = iDays + CInt(aiDays(iIndex))
											dAmounts11 = dAmounts11 + CDbl(adAmounts11(iIndex))
										Next
										iDays = iDays + CInt(aiDays(iIndex))
										dAmounts11 = dAmounts11 + CDbl(adAmounts11(iIndex))
										If iDays > 360 Then iDays = 360
										sRowContents = Replace(sRowContents, "<AMOUNT_11 />", FormatNumber(FormatNumber((dAmounts11 / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
									Else
										For iIndex = 0 To UBound(aiDays) - 1
											sRowContents = Replace(sRowContents, "<AMOUNT_11 />", FormatNumber(FormatNumber((adAmounts11(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
										Next
									End If
									Call SaveTextToFile(Server.MapPath(sFileName & ".xls"), sRowContents, "")
								End If
								lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								dAmounts = CLng(oRecordset.Fields("TotalAmount").Value)
								dAmounts3 = CLng(oRecordset.Fields("TotalAmount3").Value)
								dAmounts11 = CLng(oRecordset.Fields("TotalAmount11").Value)
								aiDays = ""
								adAmounts = ""
								adAmounts3 = ""
								adAmounts11 = ""
								bContinuous = True
								bContinuous3 = True
								bContinuous11 = True
								lTempStartDate = lTempEndDate
							End If
							bContinuous = (bContinuous And (lTempStartDate = lTempEndDate) And (dAmounts <= CLng(oRecordset.Fields("TotalAmount").Value)))
							bContinuous3 = (bContinuous And (lTempStartDate = lTempEndDate) And (dAmounts3 <= CLng(oRecordset.Fields("TotalAmount3").Value)))
							bContinuous11 = (bContinuous And (lTempStartDate = lTempEndDate) And (dAmounts11 <= CLng(oRecordset.Fields("TotalAmount11").Value)))
							lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
							dAmounts = CLng(oRecordset.Fields("TotalAmount").Value)
							dAmounts3 = CLng(oRecordset.Fields("TotalAmount3").Value)
							dAmounts11 = CLng(oRecordset.Fields("TotalAmount11").Value)

							sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
							Err.Clear
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp)
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value), -1, -1, -1)
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelID").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("Days1").Value)
							sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("Days3").Value)
							sRowContents = sRowContents & TABLE_SEPARATOR & CInt(oRecordset.Fields("Days1").Value) - CInt(oRecordset.Fields("Days3").Value)
							aiDays = aiDays & (CInt(oRecordset.Fields("Days1").Value) - CInt(oRecordset.Fields("Days3").Value)) & LIST_SEPARATOR
							sRowContents = sRowContents & TABLE_SEPARATOR & "<DAYS />"
							adAmounts = adAmounts & CDbl(oRecordset.Fields("TotalAmount").Value) & LIST_SEPARATOR
							adAmounts3 = adAmounts3 & CDbl(oRecordset.Fields("TotalAmount3").Value) & LIST_SEPARATOR
							adAmounts11 = adAmounts11 & CDbl(oRecordset.Fields("TotalAmount11").Value) & LIST_SEPARATOR
							sRowContents = sRowContents & TABLE_SEPARATOR & "<AMOUNT />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<AMOUNT />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<AMOUNT_3 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<AMOUNT_11 />"
							'sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) / 15 * 40), 2, True, False, True)
							'sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) / 15 * 40 * ((CInt(oRecordset.Fields("Days1").Value) - CInt(oRecordset.Fields("Days3").Value)) / 360)), 2, True, False, True)

							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".xls"), GetTableRowText(asRowContents, True, ""), sErrorDescription)
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						oRecordset.Close
						aiDays = Split(aiDays, LIST_SEPARATOR)
						adAmounts = Split(adAmounts, LIST_SEPARATOR)
						adAmounts3 = Split(adAmounts3, LIST_SEPARATOR)
						adAmounts11 = Split(adAmounts11, LIST_SEPARATOR)
						sRowContents = GetFileContents(Server.MapPath(sFileName & ".xls"), sErrorDescription)
						If bContinuous Then
							iDays = 0
							dAmounts = 0
							For iIndex = 0 To UBound(aiDays) - 2
								sRowContents = Replace(sRowContents, "<DAYS />", "0", 1, 1, vbBinaryCompare)
								sRowContents = Replace(sRowContents, "<AMOUNT />", "0.00", 1, 1, vbBinaryCompare)
								sRowContents = Replace(sRowContents, "<AMOUNT />", "0.00", 1, 1, vbBinaryCompare)
								iDays = iDays + CInt(aiDays(iIndex))
								dAmounts = dAmounts + CDbl(adAmounts(iIndex))
							Next
							iDays = iDays + CInt(aiDays(iIndex))
							dAmounts = dAmounts + CDbl(adAmounts(iIndex))
							If iDays > 360 Then iDays = 360
							sRowContents = Replace(sRowContents, "<DAYS />", iDays, 1, 1, vbBinaryCompare)
							sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((dAmounts / 15 * 40), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
							sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((dAmounts / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
						Else
							For iIndex = 0 To UBound(aiDays) - 1
								sRowContents = Replace(sRowContents, "<DAYS />", aiDays(iIndex), 1, 1, vbBinaryCompare)
								sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((adAmounts(iIndex) / 15 * 40), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
								sRowContents = Replace(sRowContents, "<AMOUNT />", FormatNumber(FormatNumber((adAmounts(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
							Next
						End If
						If bContinuous3 Then
							iDays = 0
							dAmounts3 = 0
							For iIndex = 0 To UBound(aiDays) - 2
								sRowContents = Replace(sRowContents, "<AMOUNT_3 />", "0.00", 1, 1, vbBinaryCompare)
								iDays = iDays + CInt(aiDays(iIndex))
								dAmounts3 = dAmounts3 + CDbl(adAmounts3(iIndex))
							Next
							iDays = iDays + CInt(aiDays(iIndex))
							dAmounts3 = dAmounts3 + CDbl(adAmounts3(iIndex))
							If iDays > 360 Then iDays = 360
							sRowContents = Replace(sRowContents, "<AMOUNT_3 />", FormatNumber(FormatNumber((dAmounts3 / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
						Else
							For iIndex = 0 To UBound(aiDays) - 1
								sRowContents = Replace(sRowContents, "<AMOUNT_3 />", FormatNumber(FormatNumber((adAmounts3(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
							Next
						End If
						If bContinuous11 Then
							iDays = 0
							dAmounts11 = 0
							For iIndex = 0 To UBound(aiDays) - 2
								sRowContents = Replace(sRowContents, "<AMOUNT_11 />", "0.00", 1, 1, vbBinaryCompare)
								iDays = iDays + CInt(aiDays(iIndex))
								dAmounts11 = dAmounts11 + CDbl(adAmounts11(iIndex))
							Next
							iDays = iDays + CInt(aiDays(iIndex))
							dAmounts11 = dAmounts11 + CDbl(adAmounts11(iIndex))
							If iDays > 360 Then iDays = 360
							sRowContents = Replace(sRowContents, "<AMOUNT_11 />", FormatNumber(FormatNumber((dAmounts11 / 15 * 40 * (iDays / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
						Else
							For iIndex = 0 To UBound(aiDays) - 1
								sRowContents = Replace(sRowContents, "<AMOUNT_11 />", FormatNumber(FormatNumber((adAmounts11(iIndex) / 15 * 40 * (aiDays(iIndex) / 360)), 2, True, False, True), 2, True, False, True), 1, 1, vbBinaryCompare)
							Next
						End If
						Call SaveTextToFile(Server.MapPath(sFileName & ".xls"), sRowContents, "")
					End If
				lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".xls"), "</TABLE>", sErrorDescription)


				lErrorNumber = ZipFile(Server.MapPath(sFileName & ".xls"), Server.MapPath(sFileName & ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(Server.MapPath(sFileName & ".xls"), sErrorDescription)
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
	BuildReport1101 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1102(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de tabuladores
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1102"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Dim lCurrentPaymentCenterID
	Dim asStateNames
	Dim asConceptNames
	Dim asPath
	Dim sZoneName

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If Len(sCondition) > 0 Then
		sCondition = Replace(sCondition, "XXX", "Missing")
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.PaymentCenterID, Employees.EmployeeNumber, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, Concepts.ConceptName, EmployeesAdjustmentsLKP.ConceptAmount, EmployeesAdjustmentsLKP.MissingDate, EmployeesAdjustmentsLKP.PaymentDate, EmployeesAdjustmentsLKP.PayrollDate, EmployeesAdjustmentsLKP.ModifyDate, JobNumber, UserName + ' ' + UserLastName As UserFullName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, CompanyShortName, CompanyName From Concepts, Employees, EmployeesAdjustmentsLKP, Jobs, Users, Areas, Areas As PaymentCenters, Zones As AreasZones, Zones As ParentZones, Zones, Companies Where (EmployeesAdjustmentsLKP.EmployeeID = Employees.EmployeeID)	And (EmployeesAdjustmentsLKP.ConceptID = Concepts.ConceptID) And (EmployeesAdjustmentsLKP.UserID = Users.UserID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) " & sCondition, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Zona,Clave CT,Nombre CT,No de empleado,Nombre,Concepto,Monto,Fecha de omisión de pago,Quincena de aplicación,Fecha de registro,Usuario que registro", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
					If Len(asPath(2)) > 0 Then
						sZoneName = CStr(asStateNames(CInt(asPath(2))))
					Else
						sZoneName = "Ninguna"
					End If
					sRowContents = CleanStringForHTML(sZoneName)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("MissingDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("PayrollDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value))
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
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1102 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1103(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de movimientos
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1103"
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
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If Len(sCondition) > 0 Then
		sCondition = Replace(sCondition, "XXX", "EmployeesHistoryList.Modify")
	End If

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeNumber, ReasonShortName, ReasonName, Jobs.JobNumber, Positions.PositionShortName, Levels.LevelShortName, Employees.RFC, Employees.CURP, Employees.SocialSecurityNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, GenderName, MaritalStatusName, UserName, UserLastName, EmployeesHistoryList.ModifyDate, Areas.AreaCode, Areas.AreaName From Areas, Employees, EmployeesChangesLKP, EmployeesHistoryList, Genders, Jobs, Levels, MaritalStatus, Positions, Reasons, Users Where (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (Employees.GenderID=Genders.GenderID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.UserID=Users.UserID) And (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (EmployeesHistoryList.EmployeeDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") " & sCondition & " Order By UserName", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1103.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ReasonShortName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("JobNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("RFC").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></B></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PositionShortName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("LevelShortName").Value) & "</FONT></B></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ReasonName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UserLastName").Value) & " " & CStr(oRecordset.Fields("UserName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value)) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></B></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("CURP").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("SocialSecurityNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("MaritalStatusName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("GenderName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""5"" ALIGN=""CENTER"">&nbsp;</TD>"
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1103 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1104(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de movimientos por usuario
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1104"
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
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition
	Dim lConceptID

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	
	If Len(sCondition) > 0 Then 
		sCondition = Replace(sCondition, "XXXDate", "EmployeesHistoryList.ModifyDate")
		If Len(lPayrollID) > 0 Then
			If Len(oRequest("ReasonID").Item) = 0 Then
				sCondition = sCondition & " And ((EmployeesHistoryList.PayrollDate=" & lPayrollID & ") Or ((EmployeesHistoryList.PayrollDate = 0) And (EmployeesHistoryList.ReasonID IN (54,55,57))))"
			Else
				sCondition = sCondition & " And (EmployeesHistoryList.PayrollDate=" & lPayrollID & ")"
			End If
		End If
	End If
	If CInt(oRequest("ReasonID").Item) = 51 Then
		sCondition = Replace(sCondition,"ReasonID In (51)","ReasonID In (51,54,55)")
	End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeNumber, ReasonShortName, StatusEmployees.StatusName, ReasonName, Jobs.JobNumber, Positions.PositionShortName, Levels.LevelShortName, GGL.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, EmployeesHistoryList.ClassificationID, Employees.RFC, Employees.CURP, Employees.SocialSecurityNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, GenderName, MaritalStatusName, UserName, UserLastName, EmployeesHistoryList.ModifyDate, Areas.AreaCode, Areas.AreaName, Jobs.ServiceID, Services.ServiceShortName, Services.ServiceName, Journeys.JourneyShortName, Journeys.JourneyName, Employees.StartHour3, Employees.EndHour3, EmployeesHistoryList.RiskLevel, EmployeesHistoryList.PositionTypeID, EmployeesHistoryList.EmployeeTypeID, ShiftName  From Areas, Employees, EmployeesHistoryList, Genders, Jobs, Levels, MaritalStatus, Positions, Reasons, StatusEmployees, Users, Services, Journeys, Shifts, GroupGradeLevels GGL  Where (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.StatusID = StatusEmployees.StatusID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID = GGL.GroupGradeLevelID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (Employees.GenderID=Genders.GenderID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.UserID=Users.UserID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.ServiceID = Services.ServiceID) And (EmployeesHistoryList.ShiftID = Shifts.ShiftID) And (Jobs.JourneyID = Journeys.JourneyID) " & sCondition & " Order By UserName Asc,  EmployeesHistoryList.EmployeeID Asc, EmployeesHistoryList.EmployeeDate Desc, EmployeesHistoryList.EndDate Desc", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
				sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".htm"
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1104.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
'				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("ReasonShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""5"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""></FONT></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("JobNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("RFC").Value) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD COLSPAN=""2"" ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD COLSPAN=""2"" ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></TD>"
						If InStr(1, ",340,341,342,343,344,345,346,348,346,349,350,", "," & oRecordset.Fields("ReasonShortName").Value & ",", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & "</FONT></TD>"
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">Indefinida</FONT></TD>"
							ElseIf CLng(oRecordset.Fields("EndDate").Value) = 0 Then
								sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> --- </FONT></TD>"
							Else
								sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
							End If
						Else
							sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> --- </FONT></TD>"
							sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & DisplayNumericDateFromSerialNumber(AddDaysToSerialDate(CLng(oRecordset.Fields("EmployeeDate").Value), -1)) & "</FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("PositionShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("LevelShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & "</FONT></TD>"
						If oRecordset.Fields("ClassificationID").Value = -1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> --- </FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("ClassificationID").Value) & "</FONT></TD>"
						End If
						If oRecordset.Fields("IntegrationID").Value = -1 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> --- </FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("IntegrationID").Value) & "</FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">&nbsp;</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("ReasonName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & "(" & CStr(oRecordset.Fields("ServiceShortName").Value) & ") " & CStr(oRecordset.Fields("ServiceName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & "(" & CStr(oRecordset.Fields("JourneyShortName").Value) & ") " &  CStr(oRecordset.Fields("JourneyName").Value) & " " & CStr(oRecordset.Fields("ShiftName").Value) &"</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("AreaName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("UserLastName").Value) & " " & CStr(oRecordset.Fields("UserName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value)) & "</FONT></TD>"
						lConceptID = CheckFor0708Concepts(oADODBConnection, CLng(oRecordset.Fields("EmployeeNumber").Value))
						If lConceptID > 0 Then
							If CLng(oRecordset.Fields("StartHour3").Value) <> CLng(oRecordset.Fields("EndHour3").Value) Then
								If lConceptID = 7 Then
									If CLng(oRecordset.Fields("StartHour3").Value) > 0 Then
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & Left(CStr(oRecordset.Fields("StartHour3").Value), 2) & ":" & Right(CStr(oRecordset.Fields("StartHour3").Value), 2) & " a " & Left(CStr(oRecordset.Fields("EndHour3").Value), 2) & ":" & Right(CStr(oRecordset.Fields("EndHour3").Value), 2) & "</FONT></TD>"
									Else
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> * </FONT></TD>"
									End If
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
								Else
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
									If CLng(oRecordset.Fields("StartHour3").Value) > 0 Then
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & Left(CStr(oRecordset.Fields("StartHour3").Value), 2) & ":" & Right(CStr(oRecordset.Fields("StartHour3").Value), 2) & " a " & Left(CStr(oRecordset.Fields("EndHour3").Value), 2) & ":" & Right(CStr(oRecordset.Fields("EndHour3").Value), 2) & "</FONT></TD>"
									Else
										sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> * </FONT></TD>"
									End If
								End If
							Else
								If lConceptID = 7 Then
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> * </FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
								ElseIf lConceptID = 8 Then
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
									sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> * </FONT></TD>"
								End If
							End If
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
						End If
						If CInt(oRecordset.Fields("RiskLevel").Value) >= 10 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("RiskLevel").Value) & "</FONT></TD>"
						ElseIf CInt(oRecordset.Fields("RiskLevel").Value) >= 1 And CInt(oRecordset.Fields("RiskLevel").Value) < 10 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("RiskLevel").Value * 10) & "</FONT></TD>"
						ElseIf CInt(oRecordset.Fields("RiskLevel").Value) = 0 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1""> &nbsp; </FONT></TD>"
						End If
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("CURP").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("SocialSecurityNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("MaritalStatusName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("GenderName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""7"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""1"">" & CStr(oRecordset.Fields("StatusName").Value) & "</FONT></TD>"
						
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1104 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1105(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de movimientos
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1105"
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
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If Len(sCondition) > 0 Then 
		sCondition = Replace(sCondition, "XXXDate", "EmployeesHistoryList.ModifyDate")
	End If

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeNumber, ReasonShortName, ReasonName, Jobs.JobNumber, Positions.PositionShortName, Levels.LevelShortName, Employees.RFC, Employees.CURP, Employees.SocialSecurityNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, GenderName, MaritalStatusName, UserName, UserLastName, EmployeesHistoryList.ModifyDate, Areas.AreaCode, Areas.AreaName From Employees, EmployeesHistoryList, Genders, Jobs, Areas, Zones, Levels, MaritalStatus, Positions, Reasons, Users Where (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (Employees.GenderID=Genders.GenderID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.UserID=Users.UserID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PayrollDate=" & lPayrollID & ") " & sCondition & " Order By UserName", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1103.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ReasonShortName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("JobNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("RFC").Value) & "</FONT></B></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD COLSPAN=""2"" ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></B></TD>"
						Else
							sRowContents = sRowContents & "<TD COLSPAN=""2"" ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></B></TD>"
						End If
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & "</FONT></B></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></B></TD>"
						Else
							If CLng(oRecordset.Fields("EndDate").Value) <> 0 Then
								sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></B></TD>"
							Else
								sRowContents = sRowContents & "<TD ROWSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""> --- </FONT></B></TD>"
							End If
						End If
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PositionShortName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("LevelShortName").Value) & "</FONT></B></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ReasonName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""3"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UserLastName").Value) & " " & CStr(oRecordset.Fields("UserName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value)) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2""></FONT></B></TD>"
					sRowContents = sRowContents & "</TR>"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("CURP").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("SocialSecurityNumber").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("MaritalStatusName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""2"" ALIGN=""CENTER""><B><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("GenderName").Value) & "</FONT></B></TD>"
						sRowContents = sRowContents & "<TD COLSPAN=""5"" ALIGN=""CENTER"">&nbsp;</TD>"
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1105 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1106(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal de honorarios
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1106"
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
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordDate, ConceptAmount, AreaCode, AreaName, ZoneName, Employees.EmployeeNumber, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesChangesLKP.PayrollDate From Areas, Employees, EmployeesChangesLKP, EmployeesHistoryList, Jobs, Payroll_" & lPayrollID & " As Sueldo, Zones Where (Employees.EmployeeID=Sueldo.EmployeeID) And (Sueldo.ConceptID=1) And (Employees.JobID=Jobs.JobID) And (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Employees.EmployeeTypeID=7) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ")" & sCondition, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1106.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(CLng(Right(GetSerialNumberForDate(""), Len("000000")))))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CLng(oRecordset.Fields("RecordDate").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaCode").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ZoneName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AreaName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1106 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1107sp(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal de que tiene reclamos registrados en EmployeesAdjustmentsLKP
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1107sp"
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

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	
	'If (InStr(1, oRequest, "StartDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartDate", "EndDate", "XXXDate", False, sCondition2)
	'sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	'If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesConceptsLKP.Start") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesConceptsLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesConceptsLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesConceptsLKP.Start", 1, 1, vbBinaryCompare) & "))"
	'sCondition = sCondition & " And EmployeesConceptsLKP.ConceptID In (" & BENEFIT_CONCEPTS_FOR_PAYROLL & ") "

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName, ConceptShortName, ConceptName, ConceptAmount, Users.UserLastName + ' ' + Users.UserName As UserFullName, MissingDate, PayrollDate From Employees, EmployeesAdjustmentsLKP, Concepts, Users Where Employees.EmployeeID=EmployeesAdjustmentsLKP.EmployeeID And Concepts.ConceptID=EmployeesAdjustmentsLKP.ConceptID And Users.UserID=EmployeesAdjustmentsLKP.UserID " & sCondition & sCondition2, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave concepto</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de omisión del pago</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Monto del reclamo</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">Nombre del beneficiario</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER"">&nbsp;</TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ConceptShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ConceptName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("QttyName").Value) & "</FONT></TD>"
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1107sp = lErrorNumber
	Err.Clear
End Function

Function BuildReport1108(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de personal de que tiene conceptos registrado en EmployeesConceptsLKP
'         Jefatura de Servicios de Personal
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1108"
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

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asConceptNames
	Dim asPath
	Dim iCount
	Dim aiConceptTotals
	Dim aiConceptGrandTotals
	Dim iIndex
	Dim sConceptShortName
	Dim sConceptStatus
	Dim bFirst
	Dim lTotal
	Dim iMin
	Dim iMax

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")

	If (InStr(1, oRequest, "StartDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "EndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartDate", "EndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "EmployeesConceptsLKP.Start") & ") Or (" & Replace(sCondition2, "XXX", "EmployeesConceptsLKP.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "EmployeesConceptsLKP.End", 1, 1, vbBinaryCompare), "XXX", "EmployeesConceptsLKP.Start", 1, 1, vbBinaryCompare) & "))"
	sCondition = sCondition & " And EmployeesConceptsLKP.ConceptID In (" & EMPLOYEES_CONCEPTS & ") "

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(ConceptID) As Max From Concepts", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value)
		End If
	End If
	For iMin = 0 To iMax
		asConceptNames = asConceptNames & LIST_SEPARATOR & ""
		aiConceptTotals = aiConceptTotals & LIST_SEPARATOR & "0"
		aiConceptGrandTotals = aiConceptGrandTotals & LIST_SEPARATOR & "0"
	Next
	asConceptNames = Split(asConceptNames, LIST_SEPARATOR)
	aiConceptTotals = Split(aiConceptTotals, LIST_SEPARATOR)
	aiConceptGrandTotals = Split(aiConceptGrandTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiConceptTotals(iIndex) = 0
		aiConceptGrandTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptShortName From Concepts Where (ConceptID>0) Order By ConceptID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		iCount=0
		Do While Not oRecordset.EOF
			asConceptNames(CInt(oRecordset.Fields("ConceptID").Value)) = SizeText(CStr(CleanStringForHTML(oRecordset.Fields("ConceptShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.PaymentCenterID, Employees.PaymentCenterID, Employees.EmployeeNumber, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, ConceptShortName, ConceptName, EmployeesConceptsLKP.ConceptID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.Active, Comments, ConceptAmount, QttyName, RegistrationDate, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath From Concepts, Employees, EmployeesConceptsLKP, QttyValues, Users, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.ConceptQttyID=QttyValues.QttyID) And (EmployeesConceptsLKP.EmployeeID=Employees.EmployeeID) And (EmployeesConceptsLKP.StartUserID=Users.UserID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) " & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & "Select Employees.EmployeeID, Employees.PaymentCenterID, Employees.EmployeeNumber, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, ConceptShortName, ConceptName, EmployeesConceptsLKP.ConceptID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, EmployeesConceptsLKP.Active, Comments, ConceptAmount, QttyName, RegistrationDate, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath From Concepts, Employees, EmployeesConceptsLKP, QttyValues, Users, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.ConceptQttyID=QttyValues.QttyID) And (EmployeesConceptsLKP.EmployeeID=Employees.EmployeeID) And (EmployeesConceptsLKP.StartUserID=Users.UserID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID = Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=ParentZones.ZoneID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) " & sCondition & sCondition2 & " Order By PaymentCenters.ParentID, PaymentCenters.AreaCode, Employees.EmployeeID" & " -->" & vbNewLine
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
							sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
							sRowContents = sRowContents & "<TD>TOTAL</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							For iIndex = 0 To UBound(aiConceptTotals)
								lTotal = CInt(aiConceptTotals(iIndex))
								If lTotal > 0 Then
									sConceptShortName = Trim(asConceptNames(CInt(iIndex)))
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & sConceptShortName & "</TD>"
										sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
									sRowContents = sRowContents & "</FONT></TR>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							Next
							For iIndex = 0 To UBound(aiConceptTotals)
								aiConceptTotals(iIndex) = 0
							Next
							sRowContents = "</TABLE>"
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
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave del entro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del centro de trabajo</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave concepto</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Descripción</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de termino</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">Cantidad</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">Moneda/%</FONT></TD>"
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">Estatus</FONT></TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeFullName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PaymentCenterShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("PaymentCenterName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ConceptName").Value) & "</FONT></TD>"
						If CInt(oRecordset.Fields("ConceptName").Value) = 93 Then
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
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & sNightShiftsDatesDesc1 & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						End If
						If CInt(oRecordset.Fields("ConceptName").Value) = 93 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>" 
						Else
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">A la fecha</FONT></TD>"
							ElseIf CLng(oRecordset.Fields("EndDate").Value) = 0 Then
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>" 
							Else
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
							End If
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value)) & "</FONT></TD>"
						If CInt(oRecordset.Fields("ConceptName").Value) = 93 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>" 
						Else
							sRowContents = sRowContents & "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & "</FONT></TD>"
						End If
						If CInt(oRecordset.Fields("ConceptName").Value) = 93 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">NA</FONT></TD>" 
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("QttyName").Value) & "</FONT></TD>"
						End If
						Select Case CInt(oRecordset.Fields("Active").Value)
							Case 0
								sConceptStatus = "En proceso"
							Case 1
								sConceptStatus = "Activo"
							Case 2
								sConceptStatus = "Cancelado"
						End Select
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & sConceptStatus & "</FONT></TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					aiConceptTotals(CInt(oRecordset.Fields("ConceptID").Value)) = aiConceptTotals(CInt(oRecordset.Fields("ConceptID").Value)) + 1
					aiConceptGrandTotals(CInt(oRecordset.Fields("ConceptID").Value)) = aiConceptGrandTotals(CInt(oRecordset.Fields("ConceptID").Value)) + 1
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
					For iIndex = 0 To UBound(aiConceptTotals)
						lTotal = CInt(aiConceptTotals(iIndex))
						If lTotal > 0 Then
							sConceptShortName = Trim(asConceptNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sConceptShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
					For iIndex = 0 To UBound(aiConceptTotals)
						aiConceptTotals(iIndex) = 0
					Next
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				End If

				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>TOTALES DEL REPORTE</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<TABLE WIDTH=""100%"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR><FONT FACE=""Arial"" SIZE=""2"">"
				sRowContents = sRowContents & "<TD>CLAVE DE INCIDENCIA</TD>"
				sRowContents = sRowContents & "<TD>TOTAL</TD>"
				sRowContents = sRowContents & "</FONT></TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				For iIndex = 0 To UBound(aiConceptGrandTotals)
					lTotal = CInt(aiConceptGrandTotals(iIndex))
					If lTotal > 0 Then
						sConceptShortName = Trim(asConceptNames(CInt(iIndex)))
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & sConceptShortName & "</TD>"
							sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				Next
				For iIndex = 0 To UBound(aiConceptTotals)
					aiConceptTotals(iIndex) = 0
				Next
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
				
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
			oZonesRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1108 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1109(oRequest, oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: Reporte de FM1
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1109"
	Dim bEmpty
	Dim dConceptAmount
	Dim oRecordset
	Dim oEmployeeRecordset
	Dim oEmployeeConceptRecordset
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeContents
	Dim asEmployeeContents
	Dim sFields
	Dim sFileEmployees
	Dim sFilePath
	Dim sFileName
	Dim sHeaderContents
	Dim sHour
	Dim sQuery
	Dim sTables
	Dim lClassificationID
	Dim lEconomicZoneID
	Dim lEmployeeDate
	Dim lEmployeeTypeID
	Dim lPositionID
	Dim lErrorNumber
	Dim lGroupGradeLevelID
	Dim lIntegrationID
	Dim lPositionTypeID
	Dim lJobID
	Dim lLevelID
	Dim lPayrollNumber
	Dim lReasonID
	Dim lReportID
	Dim lTotalChildren
	Dim sPerception
	Dim oReasonRecordset

	If Not bForExport Then
		Call GetConditionFromURL(oRequest, sCondition, -1, -1)
		If Len(sCondition) > 0 Then
			sCondition = Replace(sCondition, "XXX", "EmployeesHistoryList.Employee")
		End If
	Else
		sCondition = "And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	End If

	sFields = "Areas.AreaCode, Areas.AreaName, Areas.EconomicZoneID, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.ReasonID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.LevelID, EmployeesHistoryList.GroupGradeLevelID, EmployeesHistoryList.IntegrationID, EmployeesHistoryList.ClassificationID, EmployeesHistoryList.WorkingHours, EmployeesHistoryList.JobID, EmployeesHistoryList.PositionTypeID, EmployeeLastName, EmployeeLastName2, EmployeeName, RFC, CURP, BirthDate, Reasons.ReasonShortName, Reasons.ReasonName, Journeys.JourneyShortName, Genders.GenderName, MaritalStatus.MaritalStatusName, Employees.SocialSecurityNumber, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.PayrollDate, EmployeesHistoryList.EmployeeNumber, EmployeesHistoryList.Comments, CompanyShortName, CompanyName, Jobs.JobNumber, Jobs.StartDate As JobStartDate, Jobs.EndDate As JobEndDate, StateZones.ZoneName As StateName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EmployeesHistoryList.JobID As JobNumber, PositionShortName, PositionName, LevelShortName, Services.ServiceShortName, Services.ServiceName, StatusJobs.StatusShortName As JobStatusShortName, EmployeesHistoryList.RiskLevel, EmployeesHistoryList.PositionTypeID, JobTypeName, Positions.PositionID, PositionTypeName, PositionTypeShortName, GroupGradeLevels.GroupGradeLevelShortName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2"
	sTables = "Areas, Companies, Employees, EmployeesHistoryList, Genders, GroupGradeLevels, Jobs, JobTypes, Journeys, Levels, MaritalStatus, Areas As PaymentCenters, Positions, PositionTypes, Reasons, Services, Shifts, Zones, Zones As ParentZones, Zones As StateZones, StatusJobs"
	If Not bForExport Then
		'sCondition = "(EmployeesHistoryList.EmployeeTypeID>-1) And (EmployeesHistoryList.EmployeeTypeID<7) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (Zones.ParentID=ParentZones.ZoneID) And (ParentZones.ParentID=StateZones.ZoneID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=StatusJobs.StatusID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (bProcessed<>1) And (EmployeesHistoryList.ReasonID <> 0) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) " & sCondition
		sCondition = "(EmployeesHistoryList.EmployeeTypeID>-1) And (EmployeesHistoryList.EmployeeTypeID<7) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (Zones.ParentID=ParentZones.ZoneID) And (ParentZones.ParentID=StateZones.ZoneID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=StatusJobs.StatusID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.ReasonID <> 0) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) " & sCondition
	Else
		sCondition = "(EmployeesHistoryList.EmployeeTypeID>-1) And (EmployeesHistoryList.EmployeeTypeID<7) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (Zones.ParentID=ParentZones.ZoneID) And (ParentZones.ParentID=StateZones.ZoneID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=StatusJobs.StatusID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.ReasonID <> 0) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) " & sCondition
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFilePath = sFilePath & "\"
		sFileEmployees = sFilePath & "Employees_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
		sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
		'sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
		If Not bForExport Then
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
		End If
		sHeaderContents = " "
		oStartDate = Now()
		bEmpty = True
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesHistoryList.EmployeeID From " & sTables & " Where " & sCondition, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: " & "Select Distinct EmployeesHistoryList.EmployeeID From " & sTables & " Where " & sCondition & " -->" & vbNewLine
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				lErrorNumber = AppendTextToFile(sFileEmployees, CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			sEmployeeContents = GetFileContents(sFileEmployees, sErrorDescription)
			asEmployeeContents = Split(sEmployeeContents, vbNewLine)
		End If
		For iIndex = 0 To UBound(asEmployeeContents)
			If Len(asEmployeeContents(iIndex)) > 0 Then
				sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
				If Not bForExport Then
					sQuery = "Select " & sFields & " From " & sTables & " Where " & sCondition & " And (EmployeesHistoryList.ReasonID Not In (53,-64,-75)) And (EmployeesHistoryList.EmployeeID=" & asEmployeeContents(iIndex) & ") Order by EmployeesHistoryList.EmployeeDate Desc"
				Else
					sQuery = "Select " & sFields & " From " & sTables & " Where " & sCondition & " And (EmployeesHistoryList.ReasonID Not In (53,-64,-75)) Order by EmployeesHistoryList.EmployeeDate Desc"
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oEmployeeRecordset)
				If lErrorNumber = 0 Then
					If Not oEmployeeRecordset.EOF Then
						Do While Not oEmployeeRecordset.EOF
							sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_" & asEmployeeContents(iIndex) & ".doc"
							bEmpty = False
							lReasonID = CLng(oEmployeeRecordset.Fields("ReasonID").Value)
							sHeaderContents = GetFileContents(Server.MapPath("Templates\EmployeeForm1.htm"), sErrorDescription)
							sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							sHeaderContents = Replace(sHeaderContents, "<POSITION_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("PositionName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<JOB_NUMBER />", Right("000000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("JobNumber").Value)),6))
							sHeaderContents = Replace(sHeaderContents, "<POSITION_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("PositionShortName").Value)))
							'If B_ISSSTE Then
							'	If CLng(oEmployeeRecordset.Fields("PositionTypeID").Value) = 1 Then
							'		sHeaderContents = Replace(sHeaderContents, "<POSITION_TYPE_SHORT_NAME />", "C")
							'	Else
							'		sHeaderContents = Replace(sHeaderContents, "<POSITION_TYPE_SHORT_NAME />", "S")
							'	End If
							'Else
								sHeaderContents = Replace(sHeaderContents, "<POSITION_TYPE_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("PositionTypeShortName").Value)))
							'End If
							If CLng(oEmployeeRecordset.Fields("LevelID").Value) <> -1 Then
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LEVEL_NAME />", LEFT(RIGHT("000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("LevelShortName").Value)),3),2))
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_SUBLEVEL_NAME />", RIGHT(CleanStringForHTML(CStr(oEmployeeRecordset.Fields("LevelShortName").Value)), 1))
							Else
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LEVEL_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("GroupGradeLevelShortName").Value)))
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_SUBLEVEL_NAME />", "-")
							End If
							sHeaderContents = Replace(sHeaderContents, "<AREA_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("AreaName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<AREA_CODE />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("AreaCode").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_JOURNEY_ID />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("JourneyShortName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<NOMBRE />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EmployeeName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EmployeeName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EmployeeLastName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME2 />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EmployeeLastName2").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_RFC />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("RFC").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_CURP />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("CURP").Value)))
							If CLng(oEmployeeRecordset.Fields("EmployeeID").Value) < 1000000 Then
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EmployeeNumber").Value)))
							Else
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NUMBER />", "")
							End If
							lPayrollNumber = (CLng(Mid(CStr(oEmployeeRecordset.Fields("PayrollDate").Value), Len("00000"), Len("00"))) * 2)
							If CLng(Right(CStr(oEmployeeRecordset.Fields("EmployeeDate").Value), Len("00"))) < 11 Then
								lPayrollNumber = lPayrollNumber - 1
							End If
							sHeaderContents = Replace(sHeaderContents, "<PAYROLL_NUMBER />", lPayrollNumber)
							sHeaderContents = Replace(sHeaderContents, "<PAYROLL_YEAR />", Left(CStr(oEmployeeRecordset.Fields("PayrollDate").Value), Len("0000")))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_SSN />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("SocialSecurityNumber").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_AGE />", Abs(CalculateAgeFromSerialNumber(CLng(oEmployeeRecordset.Fields("BirthDate").Value), 0)))
							sHeaderContents = Replace(sHeaderContents, "<GENDER_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("GenderName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<MARITAL_STATUS />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("MaritalStatusName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_YEAR />", Left(CStr(oEmployeeRecordset.Fields("EmployeeDate").Value), Len("0000")))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_MONTH />", Mid(CStr(oEmployeeRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00")))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_DAY />", Right(CStr(oEmployeeRecordset.Fields("EmployeeDate").Value), Len("00")))
							If (CLng(oEmployeeRecordset.Fields("EndDate").Value) = 30000000) Or (CLng(oEmployeeRecordset.Fields("EndDate").Value) = 0) Then
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_YEAR />", "99")
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_MONTH />", "99")
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_DAY />", "99")
							Else
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_YEAR />", Left(CStr(oEmployeeRecordset.Fields("EndDate").Value), Len("0000")))
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_MONTH />", Mid(CStr(oEmployeeRecordset.Fields("EndDate").Value), Len("00000"), Len("00")))
								sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_DAY />", Right(CStr(oEmployeeRecordset.Fields("EndDate").Value), Len("00")))
							End If
							If B_ISSSTE Then
								Select Case lReasonID
									Case 12, 13, 17, 68, 18, 28
										sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", "ALTA")
									Case 21, 51, 50, 26, 57, 58
										sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", "CAMBIO")
									Case 29, 30, 31, 32, 33, 34, 43, 44, 45, 46, 47, 48
										sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", "LICENCIA")
									Case 37, 38, 39, 40, 41
										sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", "LICENCIA")
									Case 1, 2, 3, 4, 5, 6, 8, 10, 62, 63
										sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", "BAJA")
								End Select
							Else
								sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("ReasonName").Value)))
							End If
							sHeaderContents = Replace(sHeaderContents, "<REASON_SHORT_NAME />", Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("ReasonShortName").Value)), Len("0000")))
							sHeaderContents = Replace(sHeaderContents, "<ECONOMIC_ZONE_ID />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EconomicZoneID").Value)))
							sHeaderContents = Replace(sHeaderContents, "<POSITION_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("PositionName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<POSITION_TYPE_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("PositionShortName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<OCCUPATION_TYPE_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("JobTypeShortName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<JOB_STATUS_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("JobStatusShortName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<SERVICE_SHORT_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("ServiceShortName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<SERVICE_NAME />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("ServiceName").Value)))
							sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_COMMENTS />", CleanStringForHTML(CStr(oEmployeeRecordset.Fields("Comments").Value)))
							sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))),-1,-1,-1))
							If False Then
								If CLng(oEmployeeRecordset.Fields("StartHour1").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_1 />", "")
								Else
									sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("StartHour1").Value)), Len("0000"))
									sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_1 />", sHour)
								End If
								If CLng(oEmployeeRecordset.Fields("EndHour1").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_1 />", "")
								Else
									If CLng(oEmployeeRecordset.Fields("StartHour2").Value) <> 0 Then
										sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EndHour1").Value)), Len("0000"))
										sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
										sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_1 />", sHour)
										If CLng(oEmployeeRecordset.Fields("StartHour2").Value) = 0 Then
											sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_2 />", "")
										Else
											sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("StartHour2").Value)), Len("0000"))
											sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
											sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_2 />", sHour)
										End If
										If CLng(oEmployeeRecordset.Fields("EndHour2").Value) = 0 Then
											sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_2 />", "")
										Else
											sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EndHour2").Value)), Len("0000"))
											sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
											sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_2 />", sHour)
										End If
									Else
										sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_1 />", "")
										sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_2 />", "")
										sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EndHour1").Value)), Len("0000"))
										sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
										sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_2 />", sHour)
									End If
								End If
							Else
								If CLng(oEmployeeRecordset.Fields("StartHour1").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_1 />", "")
								Else
									sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("StartHour1").Value)), Len("0000"))
									sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_1 />", sHour)
								End If
								If CLng(oEmployeeRecordset.Fields("EndHour1").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_1 />", "")
								Else
									sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EndHour1").Value)), Len("0000"))
									sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_1 />", sHour)
								End If
								If CLng(oEmployeeRecordset.Fields("StartHour2").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_2 />", "")
								Else
									sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("StartHour2").Value)), Len("0000"))
									sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_HOUR_2 />", sHour)
								End If
								If CLng(oEmployeeRecordset.Fields("EndHour2").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_2 />", "")
								Else
									sHour = Right("0000" & CleanStringForHTML(CStr(oEmployeeRecordset.Fields("EndHour2").Value)), Len("0000"))
									sHour = Left(sHour, Len("00")) & ":" & Right(sHour, Len("00")) 
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_HOUR_2 />", sHour)
								End If
							End If
							lLevelID = CLng(oEmployeeRecordset.Fields("LevelID").Value)
							lGroupGradeLevelID = CLng(oEmployeeRecordset.Fields("GroupGradeLevelID").Value)
							lIntegrationID = CLng(oEmployeeRecordset.Fields("IntegrationID").Value)
							lClassificationID = CLng(oEmployeeRecordset.Fields("ClassificationID").Value)
							lEmployeeTypeID = CLng(oEmployeeRecordset.Fields("EmployeeTypeID").Value)
							lEconomicZoneID = CLng(oEmployeeRecordset.Fields("EconomicZoneID").Value)
							lPositionID = CLng(oEmployeeRecordset.Fields("PositionID").Value)
							lPositionTypeID = CLng(oEmployeeRecordset.Fields("PositionTypeID").Value)
							lJobID = CLng(oEmployeeRecordset.Fields("JobID").Value)
							lEmployeeDate = CLng(oEmployeeRecordset.Fields("EmployeeDate").Value)

							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=1) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (PositionID=" & lPositionID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_01 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_01 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=2) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (EconomicZoneID=" & lEconomicZoneID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_02 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_02 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=3) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (PositionID=" & lPositionID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_03 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_03 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=36) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (EconomicZoneID=" & lEconomicZoneID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_33 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_33 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=38) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (EconomicZoneID=" & lEconomicZoneID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_35 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_35 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=49) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ") And (LevelID=" & lLevelID & ") And (GroupGradeLevelID=" & lGroupGradeLevelID & ") And (IntegrationID=" & lIntegrationID & ") And (ClassificationID=" & lClassificationID & ") And (EconomicZoneID=" & lEconomicZoneID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_48 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_48 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							If lPositionTypeID = 4 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=12) And (EndDate=30000000) And (EmployeeTypeID=-1) And (PositionTypeID=4)", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From ConceptsValues Where (ConceptID=12) And (EndDate=30000000) And (EmployeeTypeID=" & lEmployeeTypeID & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oRecordset.Fields("ConceptAmount").Value)
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_10 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_10 />", FormatNumber(dConceptAmount, 2, True, False, True))
								End If
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (ConceptID=4) And (EndDate>=" & lEmployeeDate & ") And (EmployeeID=" & asEmployeeContents(iIndex) & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oEmployeeConceptRecordset)
							If lErrorNumber = 0 Then
								If Not oEmployeeConceptRecordset.EOF Then
									If B_ISSSTE Then
										sPerception = "*"
									Else
										dConceptAmount = CDbl(oEmployeeConceptRecordset.Fields("ConceptAmount").Value)
									End If
									If CLng(oEmployeeConceptRecordset.Fields("StartDate").Value) = 0 Then
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_YEAR />", "-")
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_MONTH />", "-")
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_DAY />", "-")
									Else
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_YEAR />", Left(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("0000")))
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_MONTH />", Mid(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
										sHeaderContents = Replace(sHeaderContents, "<JOB_START_DAY />", Right(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("00")))
									End If
									If (CLng(oEmployeeConceptRecordset.Fields("EndDate").Value) = 30000000) Then
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_YEAR />", "99")
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_MONTH />", "99")
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_DAY />", "99")
									Else
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_YEAR />", Left(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("0000")))
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_MONTH />", Mid(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("00000"), Len("00")))
										sHeaderContents = Replace(sHeaderContents, "<JOB_END_DAY />", Right(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("00")))
									End If
								Else
									dConceptAmount = 0
									sPerception = ""
								End If
								If B_ISSSTE Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_04 />", sPerception)
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_04 />", dConceptAmount & "%")
								End If
								oEmployeeConceptRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where ((ConceptID=7) Or (ConceptID=8)) And (EndDate>=" & lEmployeeDate & ") And (EmployeeID=" & asEmployeeContents(iIndex) & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oEmployeeConceptRecordset)
							If lErrorNumber = 0 Then
								If Not oEmployeeConceptRecordset.EOF Then
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_08 />", "*")
								Else
									sHeaderContents = Replace(sHeaderContents, "<HAS_CONCEPT_08 />", "")
								End If
								If CLng(oEmployeeConceptRecordset.Fields("StartDate").Value) = 0 Then
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_YEAR />", "-")
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_MONTH />", "-")
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_DAY />", "-")
								Else
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_YEAR />", Left(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("0000")))
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_MONTH />", Mid(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))
									sHeaderContents = Replace(sHeaderContents, "<JOB_START_DAY />", Right(CStr(oEmployeeConceptRecordset.Fields("StartDate").Value), Len("00")))
								End If
								If (CLng(oEmployeeConceptRecordset.Fields("EndDate").Value) = 30000000) Then
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_YEAR />", "99")
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_MONTH />", "99")
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_DAY />", "99")
								Else
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_YEAR />", Left(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("0000")))
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_MONTH />", Mid(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("00000"), Len("00")))
									sHeaderContents = Replace(sHeaderContents, "<JOB_END_DAY />", Right(CStr(oEmployeeConceptRecordset.Fields("EndDate").Value), Len("00")))
								End If
								oEmployeeConceptRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select count(*) As TotalChildren From EmployeesChildrenLKP Where (EmployeeID=" & asEmployeeContents(iIndex) & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lTotalChildren = CLng(oRecordset.Fields("TotalChildren").Value)
								Else
									lTotalChildren = 0
								End If
								sHeaderContents = Replace(sHeaderContents, "<TOTAL_CHILDREN />", lTotalChildren)
								oRecordset.Close
							End If
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesExtraInfo.*, States.StateName From EmployeesExtraInfo, States Where (EmployeesExtraInfo.StateID=States.StateID) And (EmployeeID=" & asEmployeeContents(iIndex) & ")", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sHeaderContents = Replace(sHeaderContents, "<ADDRESS_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeAddress").Value)))
									sHeaderContents = Replace(sHeaderContents, "<ADDRESS_CITY />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeCity").Value)))
									If Len(CStr(oRecordset.Fields("EmployeeZipCode").Value)) > 0 Then
										sHeaderContents = Replace(sHeaderContents, "<ADDRESS_ZIP_CODE />", "C.P. " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeZipCode").Value)) & "<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
									Else
										sHeaderContents = Replace(sHeaderContents, "<ADDRESS_ZIP_CODE />", "")
									End If
								Else
									sHeaderContents = Replace(sHeaderContents, "<ADDRESS_NAME />", "")
									sHeaderContents = Replace(sHeaderContents, "<ADDRESS_CITY />", "")
									sHeaderContents = Replace(sHeaderContents, "<ADDRESS_ZIP_CODE />", "")
								End If
								oRecordset.Close
							End If
'							If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 17) Or (lReasonID = 68) Then
'								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, Employees.EmployeeNumber, Reasons.ReasonShortName, Reasons.ReasonName, Reasons.ReasonID, JobsHistoryList.JobDate, JobsHistoryList.EndDate From Employees, EmployeesHistoryList, JobsHistoryList, Reasons Where (JobsHistoryList.JobID=" & lJobID & ") And (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=JobsHistoryList.EmployeeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeeDate<=" & lEmployeeDate & ") And (JobsHistoryList.JobDate<=" & lEmployeeDate & ") And (EmployeesHistoryList.JobID=" & lJobID & ") And (EmployeesHistoryList.ReasonID Not In (12,13,17,68)) Order by JobDate Desc", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'							Else
'								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.RFC, Employees.CURP, Employees.EmployeeNumber, Reasons.ReasonShortName, Reasons.ReasonName, Reasons.ReasonID, JobsHistoryList.JobDate, JobsHistoryList.EndDate From Employees, EmployeesHistoryList, JobsHistoryList, Reasons Where (JobsHistoryList.JobID=" & lJobID & ") And (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=JobsHistoryList.EmployeeID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeeDate<=" & lEmployeeDate & ") And (JobsHistoryList.JobDate<=" & lEmployeeDate & ") And (EmployeesHistoryList.JobID=" & lJobID & ") Order by JobDate Desc", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'							End If
							sQuery = "Select EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, Employees.EmployeeNumber," & _
										"JobDate,Jobs.EndDate" & _
									" From JobsHistoryList, Jobs, Employees" & _
									" Where JobsHistoryList.JobID = " & oEmployeeRecordset.Fields("JobID").Value & _
											" And JobsHistoryList.EndDate < " & oEmployeeRecordset.Fields("EmployeeDate").Value & _
											" And JobsHistoryList.JobID = Jobs.JobID" & _
											" And JobsHistoryList.EmployeeID = Employees.EmployeeID" & _
											" Order By JobDate Desc"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)))
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)))
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_LAST_NAME2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)))
									Else
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_LAST_NAME2 />", " ")
									End If
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
									'sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)))

									If B_ISSSTE Then
										'Select Case lReasonID
										sQuery = "Select EmployeesHistoryList.ReasonID, ReasonName, ReasonShortName, EmployeeDate, EndDate" & _
												" From EmployeesHistoryList, Reasons" & _
												" Where EmployeeID = " & oEmployeeRecordset.Fields("EmployeeID").Value & _
													" And EmployeesHistoryList.ReasonID = Reasons.ReasonID" & _
													" And EmployeeDate < 20110701" & _
												" Order By EmployeeDate Desc"
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oReasonRecordset)
										If lErrorNumber = 0 Then
											If Not oReasonRecordset.EOF Then
												Select Case CInt(oReasonRecordset.Fields("ReasonID").Value)
													Case 12, 13, 17, 68, 18, 28
														sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "ALTA")
													Case 21, 51, 50, 26, 57, 58
														sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "CAMBIO")
													Case 29, 30, 31, 32, 33, 34, 43, 44, 45, 46, 47, 48
														sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "LICENCIA")
													Case 37, 38, 39, 40, 41
														sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "LICENCIA")
													Case 1, 2, 3, 4, 5, 6, 8, 10, 62, 63
														sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "BAJA")
												End Select
											End If
										End If
									Else
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", CleanStringForHTML(CStr(oReasonRecordset.Fields("ReasonName").Value)))
									End If

									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_SHORT_NAME />", Right("0000" & CleanStringForHTML(CStr(oReasonRecordset.Fields("ReasonShortName").Value)), Len("0000")))
									If (CLng(oRecordset.Fields("JobDate").Value) = 0) Then
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_YEAR />", "-")
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_MONTH />", "-")
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_DAY />", "-")
									Else
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_YEAR />", Left(CStr(oRecordset.Fields("JobDate").Value), Len("0000")))
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_MONTH />", Mid(CStr(oRecordset.Fields("JobDate").Value), Len("00000"), Len("00")))
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_DAY />", Right(CStr(oRecordset.Fields("JobDate").Value), Len("00")))
									End If
									If (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_YEAR />", "99")
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_MONTH />", "99")
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_DAY />", "99")
									Else
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_YEAR />", Left(CStr(oRecordset.Fields("EndDate").Value), Len("0000")))
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_MONTH />", Mid(CStr(oRecordset.Fields("EndDate").Value), Len("00000"), Len("00")))
										sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_DAY />", Right(CStr(oRecordset.Fields("EndDate").Value), Len("00")))
									End If
								Else
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NAME />", "")
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME />", "")
									sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME2 />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_RFC />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_CURP />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_EMPLOYEE_NUMBER />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_NAME />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_REASON_SHORT_NAME />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_YEAR />", "&nbsp;&nbsp;")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_MONTH />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_START_DAY />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_YEAR />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_MONTH />", "")
									sHeaderContents = Replace(sHeaderContents, "<PREVIOUS_JOB_END_DAY />", "")
								End If
								oRecordset.Close
							End If
							If Not bForExport Then
								lErrorNumber = AppendTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
							Else
								Response.Write sHeaderContents
								Response.Write Err.Description
								Response.Write "<BR /><BR />"
							End If
							oEmployeeRecordset.MoveNext
							If Err.Number <> 0 Then Exit Do
						Loop
						oEmployeeRecordset.Close
					End If
				End If
			End If
		Next
		If Not bForExport Then
			If Not bEmpty Then
				lErrorNumber = DeleteFile(sFileEmployees, sErrorDescription)
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
	BuildReport1109 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1110(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte del formato de honorarios
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1110"
	Dim bEmpty
	Dim oEndDate
	Dim oRecordset
	Dim oStartDate
	Dim lErrorNumber
	Dim lReportID
	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sFields
	Dim sFileName
	Dim sFilePath
	Dim sHeaderContents
	Dim sQuery
	Dim sTables

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(sCondition) > 0 Then
		sCondition = Replace(sCondition, "XXX", "EmployeesHistoryList.Employee")
	End If
	If bForExport Then
		sCondition = "And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	End If

	sFields = "EmployeesExtraInfo.*, States.StateName, Nationality, ParentAreas.AreaName As ParentAreaName, Areas.AreaCode, Areas.AreaName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeesHistoryList.EmployeeTypeID, EmployeeLastName, EmployeeLastName2, EmployeeName, RFC, CURP, BirthDate, Reasons.ReasonShortName, Reasons.ReasonName, Genders.GenderName, MaritalStatus.MaritalStatusName, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesConceptsLKP.ConceptAmount"
	sTables = "Areas, Areas As ParentAreas, Countries, Employees, EmployeesConceptsLKP, EmployeesHistoryList, Genders, MaritalStatus, Reasons, EmployeesExtraInfo, States"
	If Not bForExport Then
		sCondition = "(EmployeesExtraInfo.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryList.EmployeeTypeID=7) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (ParentAreas.AreaID=Areas.ParentID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeID=EmployeesConceptsLKP.EmployeeID) And (EmployeesConceptsLKP.ConceptID=13) And (EmployeesHistoryList.EmployeeDate=EmployeesConceptsLKP.StartDate) And (EmployeesHistoryList.EndDate=EmployeesConceptsLKP.EndDate) And (EmployeesExtraInfo.CountryID=Countries.CountryID) And (bProcessed<>1) And (Reasons.ReasonID=14) " & sCondition
	Else
		sCondition = "(EmployeesExtraInfo.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryList.EmployeeTypeID=7) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (ParentAreas.AreaID=Areas.ParentID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeID=EmployeesConceptsLKP.EmployeeID) And (EmployeesConceptsLKP.ConceptID=13) And (EmployeesHistoryList.EmployeeDate=EmployeesConceptsLKP.StartDate) And (EmployeesHistoryList.EndDate=EmployeesConceptsLKP.EndDate) And (EmployeesExtraInfo.CountryID=Countries.CountryID) " & sCondition
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	sQuery = "Select " & sFields & " From " & sTables & " Where " & sCondition
	sDate = GetSerialNumberForDate("")
	If Not bForExport Then
		sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
		sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
		lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo obtener la información de los empleados registrados en el sistema."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If Not bForExport Then
					sFilePath = sFilePath & "\"
					sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
					sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()
					oStartDate = Now()
				End If
				bEmpty = False
				sHeaderContents = ""
				Do While Not oRecordset.EOF
					sHeaderContents = GetFileContents(Server.MapPath("Templates\HonoraryEmployeeForm.htm"), sErrorDescription)
					sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
					sHeaderContents = Replace(sHeaderContents, "<PARENT_AREA_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ParentAreaName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<AREA_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<AREA_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)))
					Else
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_LAST_NAME2 />", " ")
					End If
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
					If CLng(oRecordset.Fields("EmployeeID").Value) < 1000000 Then
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
					Else
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_NUMBER />", "")
					End If
					sHeaderContents = Replace(sHeaderContents, "<DOCUMENT_1 />", CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber1").Value)))
					sHeaderContents = Replace(sHeaderContents, "<DOCUMENT_2 />", CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber2").Value)))
					sHeaderContents = Replace(sHeaderContents, "<DOCUMENT_3 />", CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber3").Value)))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_AGE />", Abs(CalculateAgeFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value), Left(GetSerialNumberForDate(""), Len("00000000")))))
					sHeaderContents = Replace(sHeaderContents, "<GENDER_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("GenderName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<MARITAL_STATUS />", CleanStringForHTML(CStr(oRecordset.Fields("MaritalStatusName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<NATIONALITY>", CleanStringForHTML(CStr(oRecordset.Fields("Nationality").Value)))
					sHeaderContents = Replace(sHeaderContents, "<ADDRESS_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeAddress").Value)))
					sHeaderContents = Replace(sHeaderContents, "<ADDRESS_CITY />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeCity").Value)))
					sHeaderContents = Replace(sHeaderContents, "<ADDRESS_ZIP_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeZipCode").Value)))
					sHeaderContents = Replace(sHeaderContents, "<STATE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<CONCEPT_AMOUNT />", FormatNumber(CLng(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_YEAR />", Left(CStr(oRecordset.Fields("EmployeeDate").Value), Len("0000")))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_MONTH />", Mid(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00")))
					sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_DAY />", Right(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00")))
					If (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Or (CLng(oRecordset.Fields("EndDate").Value) = 0) Then
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_YEAR />", "99")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_MONTH />", "99")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_DAY />", "99")
					Else
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_YEAR />", Left(CStr(oRecordset.Fields("EndDate").Value), Len("0000")))
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_MONTH />", Mid(CStr(oRecordset.Fields("EndDate").Value), Len("00000"), Len("00")))
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_DAY />", Right(CStr(oRecordset.Fields("EndDate").Value), Len("00")))
					End If
					sHeaderContents = Replace(sHeaderContents, "<REASON_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)))
					sHeaderContents = Replace(sHeaderContents, "<REASON_SHORT_NAME />", Right("0000" & CleanStringForHTML(CStr(oRecordset.Fields("ReasonShortName").Value)), Len("0000")))
					Select Case CInt(oRecordset.Fields("EmployeeActivityID").Value)
						Case 1
							sHeaderContents = Replace(sHeaderContents, "<MEDICAL_ACTIVITIES />", "X")
							sHeaderContents = Replace(sHeaderContents, "<TECHNICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<ADMINISTRATIVE_ACTIVITIES />", "")
						Case 2
							sHeaderContents = Replace(sHeaderContents, "<MEDICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<TECHNICAL_ACTIVITIES />", "X")
							sHeaderContents = Replace(sHeaderContents, "<ADMINISTRATIVE_ACTIVITIES />", "")
						Case 3
							sHeaderContents = Replace(sHeaderContents, "<MEDICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<TECHNICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<ADMINISTRATIVE_ACTIVITIES />", "X")
						Case Else
							sHeaderContents = Replace(sHeaderContents, "<MEDICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<TECHNICAL_ACTIVITIES />", "")
							sHeaderContents = Replace(sHeaderContents, "<ADMINISTRATIVE_ACTIVITIES />", "")
					End Select
					If Not bForExport Then
						lErrorNumber = AppendTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
					Else
						Response.Write sHeaderContents
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			End If
		End If
		If Not bEmpty Then
			If Not bForExport Then
				lErrorNumber = DeleteFile(sFileEmployees, sErrorDescription)
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
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1110 = lErrorNumber
	Err.Clear
End Function
%>