<%
Function BuildReport1111(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the history list for the selected jobs
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1111"
	Dim lCurrentID
	Dim sCondition
	Dim oRecordset
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
	sCondition = Replace(Replace(Replace(sCondition, "Employees.", "JobsHistoryList."), "Jobs.", "JobsHistoryList."), ".JobNumber", ".JobID")
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And (JobsHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los historiales de las plazas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct JobsHistoryList.JobID, EmployeeID, OwnerID, CompanyShortName, CompanyName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, PositionShortName, PositionName, JobTypeShortName, JobTypeName, ShiftShortName, ShiftName, JourneyShortName, JourneyName, JobsHistoryList.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, JobsHistoryList.IntegrationID, OccupationTypeShortName, OccupationTypeName, ServiceShortName, ServiceName, LevelShortName, JobsHistoryList.WorkingHours, StatusName, JobsHistoryList.JobDate, JobsHistoryList.EndDate From JobsHistoryList, Companies, Zones, Areas, Areas As PaymentCenters, Positions, JobTypes, Shifts, Journeys, GroupGradeLevels, OccupationTypes, Services, Levels, StatusJobs Where (JobsHistoryList.CompanyID=Companies.CompanyID) And (JobsHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (JobsHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (JobsHistoryList.PositionID=Positions.PositionID) And (JobsHistoryList.JobTypeID=JobTypes.JobTypeID) And (JobsHistoryList.ShiftID=Shifts.ShiftID) And (JobsHistoryList.JourneyID=Journeys.JourneyID) And (JobsHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (JobsHistoryList.OccupationTypeID=OccupationTypes.OccupationTypeID) And (JobsHistoryList.ServiceID=Services.ServiceID) And (JobsHistoryList.LevelID=Levels.LevelID) And (JobsHistoryList.StatusID=StatusJobs.StatusID) " & sCondition & " Order By JobsHistoryList.JobID, JobDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

				lCurrentID = -2
				Do While Not oRecordset.EOF

					If lCurrentID <> CLng(oRecordset.Fields("JobID").Value) Then
						If lCurrentID <> -2 Then
							lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)
						End If
						lErrorNumber = AppendTextToFile(sFilePath, "<B>PLAZA NÚMERO " & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)) & "</B><BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
							asColumnsTitles = Split("Fecha inicio,Fecha fin,No. Emp.,No. titularidad,Compañía,Adscripción,Centro de pago,Puesto,Tipo de puesto,Horario,Turno,Clasificación,Grupo grado nivel,Integración,Tipo de ocupación,Servicio,Nivel-subnivel,Horas laboradas,Estatus", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,,,,,,CENTER,,CENTER,,,CENTER,RIGHT,", ",", -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)
						lCurrentID = CLng(oRecordset.Fields("JobID").Value)
					End If

					If CLng(oRecordset.Fields("JobDate").Value) = 0 Then
						sRowContents = "-"
					Else
						sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value), -1, -1, -1)
					End If
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					End If
					If CLng(oRecordset.Fields("EmployeeID").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & ""
					End If
					If CLng(oRecordset.Fields("OwnerID").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & Right("000000" & CStr(oRecordset.Fields("OwnerID").Value), Len("000000")) & """)"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & ""
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobTypeShortName").Value) & ". " & CStr(oRecordset.Fields("JobTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ". " & CStr(oRecordset.Fields("GroupGradeLevelName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OccupationTypeShortName").Value) & ". " & CStr(oRecordset.Fields("OccupationTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)

				lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
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
	BuildReport1111 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1112(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the history list for the selected employees
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1112"
	Dim lCurrentID
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lStartDate
	Dim lEndDate
	Dim lAntiquity
	Dim sTemp
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "Employees.", "EmployeesHistoryList.")
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los historiales de las plazas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, Employees.StartDate, CompanyShortName, CompanyName, EmployeeTypeShortName, EmployeeTypeName, EmployeesHistoryList.JobID, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, EmployeesHistoryList.EmployeeTypeID, LevelName, GroupGradeLevelName, EmployeesHistoryList.ClassificationID, EmployeesHistoryList.IntegrationID, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, ServiceShortName, ServiceName, LevelShortName, EmployeesHistoryList.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, EmployeesHistoryList.IntegrationID, JourneyShortName, JourneyName, ShiftShortName, ShiftName, EmployeesHistoryList.WorkingHours, StatusName, ReasonName, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate From Employees, EmployeesHistoryList, Companies, EmployeeTypes, Zones, Areas, Areas As PaymentCenters, Positions, PositionTypes, Shifts, Journeys, GroupGradeLevels, Services, Levels, StatusEmployees, Reasons Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeeDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			sFilePath = Server.MapPath(sFileName & ".xls")
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
				lCurrentID = -2
				lAntiquity = 0
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID <> -2 Then
'							sRowContents = "<SPAN COLS=""2"" /><B>Antigüedad total</B>" & TABLE_SEPARATOR & "<B>" & GetYearsMonthsDays(lAntiquity) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""18"" />&nbsp;"
'							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
'							lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
							lAntiquity = 0
							lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)

							lErrorNumber = AppendTextToFile(sFilePath, "<B>4. Observaciones:</B><BR />", sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath, "<B>No. de Empleado: <U>&nbsp;&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "&nbsp;&nbsp;&nbsp;&nbsp;</U> NINGUNA OBSERVACIÓN<BR />", sErrorDescription)
						End If
						lErrorNumber = AppendTextToFile(sFilePath, "<B>1. Datos del trabajador:</B><BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "Nombre Completo<BR />", sErrorDescription)
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							lErrorNumber = AppendTextToFile(sFilePath, CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "</B><BR />", sErrorDescription)
						Else
							lErrorNumber = AppendTextToFile(sFilePath, CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & vbTab & " " & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & vbTab & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "</B><BR />", sErrorDescription)
						End If
						lErrorNumber = AppendTextToFile(sFilePath, "Apellido Paterno" & vbTab & "Apellido Materno" & vbTab & "Nombre(s)" & vbTab & "R.F.C." & vbTab & "C.U.R.P." & vbTab & "<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "Domicilio Completo<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "Calle, avenida, calzada, etc" & vbTab & "No. exterior e interior" & vbTab & "Colonia, barrio o Sec." & vbTab & "C.P." & vbTab & "Ciudad" & vbTab & "Estado" & vbTab & "<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "<B>2. Periodo de aportaciones al fondo del ISSSTE:</B><BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "Fecha de ingreso" & vbTab & vbTab & "Fecha de baja" & vbTab & "<BR />", sErrorDescription)
						sTemp = Right(CStr(oRecordset.Fields("StartDate").Value), Len("00"))
						If Str(sTemp, "01", vbBinaryCompare) = 0 Then
							sTemp = "Primero"
						Else
							sTemp = FormatNumberAsText(sTemp, False)
						End If
						sTemp = sTemp & " de " & asMonthNames_es(CInt(Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")))) & " de "
						sTemp = sTemp & FormatNumberAsText(Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")), False)
						lErrorNumber = AppendTextToFile(sFilePath, DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & vbTab & UCase(sTemp) & "<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "Con número" & vbTab & "Con letra (día, mes, año)" & vbTab & "Con número" & vbTab & "Con letra (día, mes, año)<BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "<B>3. Motivo y periodo en que ocurrió la(s) baja(s), reingresos(s), y/o suspensión(es)</B><BR />", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
						asCellAlignments = Split("CENTER,,,,,,,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
						asColumnsTitles = Split("Motivo,<SPAN COLS=""6"" />Periodo,Puesto,Pagaduría,Sueldo Cotizable,Quinquenios,Otras percepciones sujetas a aportaciones al ISSSTE,Total (Pesos)", ",", -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)
						asColumnsTitles = Split("&nbsp;,Del,Al,&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)
						asColumnsTitles = Split("&nbsp;,Día,Mes,Año,Día,Mes,Año,&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, sErrorDescription), sErrorDescription)
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					End If
					If CLng(oRecordset.Fields("EmployeeDate").Value) = 0 Then
						lStartDate = -1
'						sRowContents = "-"
					Else
						lStartDate = CLng(oRecordset.Fields("EmployeeDate").Value)
'						sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value), -1, -1, -1)
					End If
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						lEndDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
'						sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
					Else
						lEndDate = CLng(oRecordset.Fields("EndDate").Value)
'						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					End If
					If lStartDate = -1 Then lStartDate = lEndDate

					sRowContents = ""
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & Right(lStartDate, Len("00")) & TABLE_SEPARATOR & Mid(lStartDate, Len("00000"), Len("00")) & TABLE_SEPARATOR & Left(lStartDate, Len("0000"))
					sRowContents = sRowContents & TABLE_SEPARATOR & Right(lEndDate, Len("00")) & TABLE_SEPARATOR & Mid(lEndDate, Len("00000"), Len("00")) & TABLE_SEPARATOR & Left(lEndDate, Len("0000"))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & " " & CStr(oRecordset.Fields("WorkingHours").Value) & " HRS."
					If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
						sRowContents = sRowContents & " (" & CleanStringForHTML("GGN: " & CStr(oRecordset.Fields("GroupGradeLevelName").Value) & ", Clas:" & CStr(oRecordset.Fields("ClassificationID").Value) & ", Int: " & CStr(oRecordset.Fields("IntegrationID").Value)) & ")"
					Else
						sRowContents = sRowContents & " (" & CleanStringForHTML("Nivel: " & CStr(oRecordset.Fields("LevelName").Value)) & ")"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & 0.00 & TABLE_SEPARATOR & 0.00 & TABLE_SEPARATOR & 0.00 & TABLE_SEPARATOR & 0.00 & TABLE_SEPARATOR & 0.00

'					sRowContents = sRowContents & TABLE_SEPARATOR & (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1)
'					lAntiquity = lAntiquity + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1)
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value) & ". " & CStr(oRecordset.Fields("PositionTypeName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ". " & CStr(oRecordset.Fields("GroupGradeLevelName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
'					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
'				sRowContents = "<SPAN COLS=""2"" /><B>Antigüedad total</B>" & TABLE_SEPARATOR & "<B>" & GetYearsMonthsDays(lAntiquity) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""18"" />&nbsp;"
'				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
'				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
				lAntiquity = 0
				lErrorNumber = AppendTextToFile(sFilePath, "</TABLE><BR /><BR />", sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "<B>4. Observaciones:</B><BR />", sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "<B>No. de Empleado: <U>&nbsp;&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(Right(("000000" & lCurrentID), Len("000000"))) & "&nbsp;&nbsp;&nbsp;&nbsp;</U> NINGUNA OBSERVACIÓN<BR />", sErrorDescription)

				lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
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
	BuildReport1112 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1115(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte del formato de honorarios
'         Carpeta 3. Arranque del servicios (anexos)
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1115"
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

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(sCondition) > 0 Then
		sCondition = Replace(sCondition, "XXX", "EmployeesHistoryList.Employee")
	End If
	If bForExport Then
		sCondition = "And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	End If

	sFields = "EmployeesExtraInfo.*, States.StateName, Nationality, ParentAreas.AreaName As ParentAreaName, Areas.AreaCode, Areas.AreaName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.ReasonID, EmployeeLastName, EmployeeLastName2, EmployeeName, RFC, CURP, BirthDate, Reasons.ReasonShortName, Reasons.ReasonName, Genders.GenderName, MaritalStatus.MaritalStatusName, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate"
	sTables = "Areas, Areas As ParentAreas, Countries, Employees, EmployeesHistoryList, Genders, MaritalStatus, Reasons, EmployeesExtraInfo, States"
	If Not bForExport Then
		sCondition = "(EmployeesExtraInfo.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryList.EmployeeTypeID=7) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (ParentAreas.AreaID=Areas.ParentID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesExtraInfo.CountryID=Countries.CountryID) And (bProcessed<>1) And (EmployeesHistoryList.ReasonID=66) " & sCondition
	Else
		sCondition = "(EmployeesExtraInfo.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryList.EmployeeTypeID=7) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (ParentAreas.AreaID=Areas.ParentID) And (Employees.GenderID=Genders.GenderID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesExtraInfo.CountryID=Countries.CountryID) " & sCondition
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
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If Not bForExport Then
					sFilePath = sFilePath & "\"
					sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
					sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()
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
					If CLng(oRecordset.Fields("ReasonID").Value) = 14 Then
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
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_YEAR />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_MONTH />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_DAY />", "")
					Else
						sHeaderContents = Replace(sHeaderContents, "<CONCEPT_AMOUNT />", "&nbsp;&nbsp;<BR /><BR />")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_YEAR />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_MONTH />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_START_DAY />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_YEAR />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_MONTH />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_END_DAY />", "")
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_YEAR />", Left(CStr(oRecordset.Fields("EmployeeDate").Value), Len("0000")))
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_MONTH />", Mid(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00")))
						sHeaderContents = Replace(sHeaderContents, "<EMPLOYEE_DROP_DAY />", Right(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00")))
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
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
						
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
				Response.Write "1" & oStartDate & "<BR />"
				Response.Write "2" & oEndDate
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
	BuildReport1115 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1116(oRequest, oADODBConnection, bForExport, lAntiquityID, sErrorDescription)
'************************************************************
'Purpose: To get the records from EmployeesHistoryList and
'         EmployeesAbsencesLKP
'Inputs:  oRequest, oADODBConnection, bForExport, lAntiquityID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1116"
	Dim oRecordset
	Dim lEmployeeID
	Dim lEmployeeStartDate
	Dim sCondition
	Dim aiYears
	Dim aiMonths
	Dim aiDays
	Dim iIndex
	Dim lDate
	Dim lTemp
	Dim bEmpty
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim lEmployeeTypeID

	aiYears = Split("0,0,0,0,0,0", ",")
	aiMonths = Split("0,0,0,0,0,0", ",")
	aiDays = Split("0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(aiYears)
		aiYears(iIndex) = 0
		aiMonths(iIndex) = 0
		aiDays(iIndex) = 0
	Next
	bEmpty = True
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If Len(sCondition) = 0 Then sCondition = " And (Employees.EmployeeID=" & oRequest("EmployeeID").Item & ")"
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	If (Len(oRequest("EmployeeYear").Item) > 0) And (Len(oRequest("EmployeeMonth").Item) > 0) And (Len(oRequest("EmployeeDay").Item) > 0) Then
		lDate = CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
	Else
		lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Employees.EmployeeTypeID, JobNumber, RFC, CURP, PositionShortName, AreaShortName From Employees, Jobs, Positions, Areas Where (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) " & sCondition, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<B>No. de empleado: " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</B><BR />"
			If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
				Response.Write "<B>Nombre del empleado: " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B><BR />"
			Else
				Response.Write "<B>Nombre del empleado: " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</B><BR />")
			End If
			Response.Write "Plaza: " & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & "<BR />"
			Response.Write "Puesto: " & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "<BR />"
			Response.Write "Centro de trabajo: " & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "<BR />"
			Response.Write "RFC: " & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "<BR />"
			Response.Write "CURP: " & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "<BR />"
			Response.Write "<B>Antigüedad hasta el día: " & DisplayDateFromSerialNumber(lDate, -1, -1,- 1) & "</B><BR />"
			Response.Write "<BR />"
			lEmployeeTypeID = oRecordset.Fields("EmployeeTypeID").Value
		End If
		oRecordset.Close
	End If

	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Fecha de inicio,Fecha de término,Años,Meses,Días,Total de días,Comentario,<SPAN COLS=""2"" />Tipo de incidencia", ",", -1, vbBinaryCompare)
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
		asCellAlignments = Split(",,RIGHT,RIGHT,RIGHT,RIGHT,,", ",", -1, vbBinaryCompare)

		sCondition = sCondition & " And (EmployeeDate<=" & lDate & ")"
		sErrorDescription = "No se pudo obtener la información de los registros."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.StatusID, StatusEmployees.Active, Reasons.ActiveEmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.Comments, StatusName, ReasonName, EmployeesHistoryList.JobID From EmployeesHistoryList, StatusEmployees, Reasons, Employees, Jobs, Areas Where (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & " Order By EmployeeDate Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				Do While Not oRecordset.EOF
					lEmployeeStartDate = CLng(oRecordset.Fields("EmployeeDate").Value)
					sRowContents = DisplayDateFromSerialNumber(lEmployeeStartDate, -1, -1, -1)
					If CLng(oRecordset.Fields("EndDate").Value) > lDate Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lDate, -1, -1, -1)
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(lEmployeeStartDate), GetDateFromSerialNumber(lDate))) + 1
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(lEmployeeStartDate), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
					End If
					aiDays(1) = aiDays(1) + lTemp
					If (CInt(oRecordset.Fields("ActiveEmployeeID").Value) = 0) Or (CLng(oRecordset.Fields("JobID").Value) = -3) Then
						aiDays(3) = aiDays(3) + lTemp
					End If
					'Call GetAntiquityFromSerialDates(lEmployeeStartDate, CLng(oRecordset.Fields("EndDate").Value), aiYears(0), aiMonths(0), aiDays(0))
					aiDays(0) = lTemp
					aiYears(0) = Int(aiDays(0) / 365)
					aiDays(0) = aiDays(0) Mod 365
					aiMonths(0) = Int(aiDays(0) / 30.4)
					aiDays(0) = Int(aiDays(0) - (aiMonths(0) * 30.4))
					sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & lTemp
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
						Err.Clear
					If (CLng(oRecordset.Fields("StatusID").Value) = -1) And (CLng(oRecordset.Fields("JobID").Value) = -2) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Laborado"
					ElseIf (CLng(oRecordset.Fields("StatusID").Value) = -1) And (CLng(oRecordset.Fields("JobID").Value) = -3) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Licencia"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value))
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
		End If
		sErrorDescription = "No se pudo obtener la información de los registros."
		If Len(oRequest("EmployeeID").Item) > 0 Then
			lEmployeeID = oRequest("EmployeeID").Item
		Else
			lEmployeeID = oRequest("EmployeeNumber").Item
		End If
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAntiquitiesLKP.*, FederalCompanyName From EmployeesAntiquitiesLKP, FederalCompanies Where (EmployeesAntiquitiesLKP.FederalCompanyID=FederalCompanies.FederalCompanyID) and (EmployeeID=" & lEmployeeID & ") Order By AntiquityDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				Do While Not oRecordset.EOF
					sRowContents = "<CENTER>---</CENTER>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("AntiquityYears").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("AntiquityMonths").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("AntiquityDays").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					aiDays(2) = aiDays(2) + CInt(oRecordset.Fields("AntiquityDays").Value)
					aiMonths(2) = aiMonths(2) + CInt(oRecordset.Fields("AntiquityMonths").Value)
					aiYears(2) = aiYears(2) + CInt(oRecordset.Fields("AntiquityYears").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""3"" />"
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("FederalCompanyName").Value))
						Err.Clear
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
		End If
		If lErrorNumber = 0 Then
			Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
		End If

		sErrorDescription = "No se pudo obtener la información de los registros."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.AbsenceID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, EmployeesAbsencesLKP.Reasons, AbsenceShortName, AbsenceName From EmployeesAbsencesLKP, Absences, Employees, Jobs, Areas Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (EmployeesAbsencesLKP.AbsenceID In (10,95)) And (EmployeesAbsencesLKP.OcurredDate<=EmployeesAbsencesLKP.EndDate) " & Replace(sCondition, "EmployeeDate", "OcurredDate") & " Order By OcurredDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.AbsenceID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, EmployeesAbsencesLKP.Reasons, AbsenceShortName, AbsenceName From EmployeesAbsencesLKP, Absences, Employees, Jobs, Areas Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (EmployeesAbsencesLKP.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,40,41,42,43,44,45,46,47,48,49)) And (EmployeesAbsencesLKP.OcurredDate<=EmployeesAbsencesLKP.EndDate) " & Replace(sCondition, "EmployeeDate", "OcurredDate") & " Order By OcurredDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False

				asRowContents = Split("<SPAN COLS=""9"" /><B>AUSENTISMOS</B>", TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				Do While Not oRecordset.EOF
					sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("EndDate").Value) > lDate Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lDate, -1, -1, -1)
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(lDate))) + 1
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
					End If
					aiDays(4) = aiDays(4) + lTemp
					'Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("OcurredDate").Value), CLng(oRecordset.Fields("EndDate").Value), aiYears(0), aiMonths(0), aiDays(0))
					aiDays(0) = lTemp
					aiYears(0) = Int(aiDays(0) / 365)
					aiDays(0) = aiDays(0) Mod 365
					aiMonths(0) = Int(aiDays(0) / 30.4)
					aiDays(0) = Int(aiDays(0) - (aiMonths(0) * 30.4))
					sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & lTemp
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
						Err.Clear
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value) & ". " & CStr(oRecordset.Fields("AbsenceName").Value))
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
		End If
	Response.Write "</TABLE><BR /><BR />"
	If bEmpty Then
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
	Else
		Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Años,Meses,Días", ",", -1, vbBinaryCompare)
			asCellWidths = Split(",,,", ",", -1, vbBinaryCompare)
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
			If (CInt(lEmployeeTypeID) <> 7) And (CInt(lEmployeeTypeID) <> 6) Then
				sRowContents = "<FONT TITLE=""" & aiDays(1) & """><B>Antigüedad ISSSTE</B></FONT>"
				aiDays(5) = aiDays(1) - aiDays(3)' - aiDays(4)
				If aiDays(5) > 0 Then
					aiYears(5) = Int(aiDays(5) / 365)
					aiDays(5) = aiDays(5) Mod 365
					aiMonths(5) = Int(aiDays(5) / 30.4)
					aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
					'Call GetAntiquityFromDays(aiDays(5), lEmployeeStartDate, aiYears(5), aiMonths(5), aiDays(5))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiYears(5) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiMonths(5) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiDays(5) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

	'			sErrorDescription = "No se pudo obtener la información de los registros."
	'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AntiquityDate, EmployeesAntiquitiesLKP.EndDate From EmployeesAntiquitiesLKP, Employees, Jobs, Areas Where (EmployeesAntiquitiesLKP.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) " & sCondition & " Order By AntiquityDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	'			If lErrorNumber = 0 Then
	'				Do While Not oRecordset.EOF
	'					aiDays(2) = aiDays(2) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("AntiquityDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
	'					oRecordset.MoveNext
	'					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
	'				Loop
	'				oRecordset.Close
	'			End If
				sRowContents = "<B>Antigüedad en otras dependencias</B>"
	'			If aiDays(2) > 0 Then Call GetAntiquityFromDays(aiDays(2), 0, aiYears(2), aiMonths(2), aiDays(2))
				If aiDays(2) >= 30 Then
					aiMonths(2) = aiMonths(2) + Int(aiDays(2) / 30)
					aiDays(2) = aiDays(2) Mod 30
				End If
				If aiMonths(2) >= 12 Then
					aiYears(2) = aiYears(2) + Int(aiMonths(2) / 12)
					aiMonths(2) = aiMonths(2) Mod 12
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(2)
				sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(2)
				sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(2)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "<B>Antigüedad Federal</B>"
				If aiDays(1) > 0 Then
					aiDays(1) = aiDays(1) - aiDays(3) - aiDays(4)
					lTemp = aiDays(1)
					aiYears(1) = Int(aiDays(1) / 365)
					aiDays(1) = aiDays(1) Mod 365
					aiMonths(1) = Int(aiDays(1) / 30.4)
					aiDays(1) = Int(aiDays(1) - (aiMonths(1) * 30.4))
					'Call GetAntiquityFromDays((aiDays(1) - aiDays(4)), lEmployeeStartDate, aiYears(1), aiMonths(1), aiDays(1))
				End If
				aiYears(1) = aiYears(1) + aiYears(2)
				aiMonths(1) = aiMonths(1) + aiMonths(2)
				aiDays(1) = aiDays(1) + aiDays(2)
				If aiDays(1) >= 30 Then
					aiDays(1) = aiDays(1) - 30
					aiMonths(1) = aiMonths(1) + 1
				End If
				If aiMonths(1) >= 12 Then
					aiMonths(1) = aiMonths(1) - 12
					aiYears(1) = aiYears(1) + 1
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(1) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(1) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(1) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				lAntiquityID = CInt(aiYears(1)) + (CInt(aiMonths(1)) / 12)
				sErrorDescription = "No se pudo obtener la información de los registros."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AntiquityID From Antiquities Where (StartYears>=" & lAntiquityID & ") Order By StartYears", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				lAntiquityID = -1
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then lAntiquityID = CInt(oRecordset.Fields("AntiquityID").Value)
					oRecordset.Close
				End If

				sRowContents = "Total de licencias"
				lTemp = aiDays(3)
				If aiDays(3) > 0 Then
					aiYears(3) = Int(aiDays(3) / 365)
					aiDays(3) = aiDays(3) Mod 365
					aiMonths(3) = Int(aiDays(3) / 30.4)
					aiDays(3) = Int(aiDays(3) - (aiMonths(3) * 30.4))
					'Call GetAntiquityFromDays(aiDays(3), lEmployeeStartDate, aiYears(3), aiMonths(3), aiDays(3))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(3) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(3) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(3) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "Total de ausentismos"
				lTemp = aiDays(4)
				If aiDays(4) > 0 Then
					aiYears(4) = Int(aiDays(4) / 365)
					aiDays(4) = aiDays(4) Mod 365
					aiMonths(4) = Int(aiDays(4) / 30.4)
					aiDays(4) = Int(aiDays(4) - (aiMonths(4) * 30.4))
					'Call GetAntiquityFromDays(aiDays(4), lEmployeeStartDate, aiYears(4), aiMonths(4), aiDays(4))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(4) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(4) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(4) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Else
				For iIndex = 0 To UBound(aiYears)
					aiYears(iIndex) = 0
					aiMonths(iIndex) = 0
					aiDays(iIndex) = 0
				Next

				sRowContents = "<FONT TITLE=""" & aiDays(1) & """><B>Antigüedad ISSSTE</B></FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiYears(5) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiMonths(5) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & aiDays(1) & """>" & aiDays(5) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = "<B>Antigüedad en otras dependencias</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(2)
				sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(2)
				sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(2)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "<B>Antigüedad Federal</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(1) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(1) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(1) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = "Total de licencias"
				lTemp = aiDays(3)
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(3) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(3) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(3) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "Total de ausentismos"
				lTemp = aiDays(4)
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiYears(4) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiMonths(4) & "</FONT>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & lTemp & """>" & aiDays(4) & "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		Response.Write "</TABLE><BR /><BR />"
	End If

	Set oRecordset = Nothing
	BuildReport1116 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1117(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To get the records from EmployeesHistoryList and
'         EmployeesAbsencesLKP
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1117"
	Const B_COMMENTS = False
	Dim oRecordset
	Dim oAbsencesRecordset
	Dim sCondition
	Dim lEmployeeStartDate
	Dim aiYears
	Dim aiMonths
	Dim aiDays
	Dim iIndex
	Dim lDate
	Dim lTemp
	Dim lEndDate
	Dim lZoneID
	Dim lAreaID
	Dim lEmployeeID
	Dim iCounter
	Dim iTotalCounter
	Dim sFileContents
	Dim asAntiquities
	Dim sResultsByIDs
	Dim bDone
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	aiYears = Split("0,0,0,0,0,0", ",")
	aiMonths = Split("0,0,0,0,0,0", ",")
	aiDays = Split("0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(aiYears)
		aiYears(iIndex) = 0
		aiMonths(iIndex) = 0
		aiDays(iIndex) = 0
	Next
	If (Len(oRequest("EmployeeYear").Item) > 0) And (Len(oRequest("EmployeeMonth").Item) > 0) And (Len(oRequest("EmployeeDay").Item) > 0) Then
		lDate = CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
	Else
		lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	End If
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees."), "PositionTypes.", "Employees."), "JobTypes.", "Jobs."), "Zones.", "Zones3.")

	asAntiquities = ""
	sResultsByIDs = ""
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, AntiquityDate, EmployeesAntiquitiesLKP.EndDate, FederalCompanies.FederalCompanyName, AntiquityYears, AntiquityMonths, AntiquityDays From Employees, EmployeesAntiquitiesLKP, Jobs, Areas, Zones As Zones3, FederalCompanies Where (Employees.EmployeeID=EmployeesAntiquitiesLKP.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (EmployeesAntiquitiesLKP.FederalCompanyID=FederalCompanies.FederalCompanyID) " & sCondition & " And (EmployeesAntiquitiesLKP.AntiquityDate<=" & lDate & ") Order By Employees.EmployeeID, AntiquityDate Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asAntiquities = asAntiquities & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AntiquityYears").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AntiquityMonths").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AntiquityDays").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("FederalCompanyName").Value) & LIST_SEPARATOR
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If
	If Len(asAntiquities) > 0 Then asAntiquities = Left(asAntiquities, (Len(asAntiquities) - Len(LIST_SEPARATOR)))
	asAntiquities = Split(asAntiquities, LIST_SEPARATOR)
	For iIndex = 0 To UBound(asAntiquities)
		asAntiquities(iIndex) = Split(asAntiquities(iIndex), SECOND_LIST_SEPARATOR)
		asAntiquities(iIndex)(0) = CLng(asAntiquities(iIndex)(0))
		asAntiquities(iIndex)(1) = CInt(asAntiquities(iIndex)(1))
		asAntiquities(iIndex)(2) = CInt(asAntiquities(iIndex)(2))
		asAntiquities(iIndex)(3) = CInt(asAntiquities(iIndex)(3))
	Next

	sCondition = sCondition & " And (EmployeeDate<=EmployeesHistoryList.EndDate)"
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Zones1.ZoneID, Zones1.ZoneCode, Zones1.ZoneName, Areas.AreaID, AreaShortName, AreaName, Employees.EmployeeID, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeDate, EmployeesHistoryList.EndDate, StatusName, StatusEmployees.Active, Reasons.ActiveEmployeeID From Employees, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones As Zones3, Zones As Zones2, Zones As Zones1 Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeeDate<=" & lDate & ") " & sCondition & " Order By Zones1.ZoneID, Areas.AreaID, Employees.EmployeeID, EmployeeDate Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No. empleado,Nombre del Empleado,RFC,CURP,Años,Meses,Días,Días antigüedad,Días licencia", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

				lZoneID = -2
				lAreaID = -2
				lEmployeeID = -2
				iCounter = 0
				iTotalCounter = 0
				Do While Not oRecordset.EOF
					If lZoneID <> CLng(oRecordset.Fields("ZoneID").Value) Then
						If lZoneID <> -2 Then
							If Not bDone Then
								For iIndex = 0 To UBound(asAntiquities)
									If (asAntiquities(iIndex)(0) = lEmployeeID) Then
										sRowContents = TABLE_SEPARATOR
										sRowContents = sRowContents & "<CENTER>---</CENTER>"
										sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Actividad en " & CleanStringForHTML(asAntiquities(iIndex)(4))
										aiYears(2) = aiYears(2) + asAntiquities(iIndex)(1)
										aiMonths(2) = aiMonths(2) + asAntiquities(iIndex)(2)
										aiDays(2) = aiDays(2) + asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(1)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(2)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & Int(asAntiquities(iIndex)(1) * 365) + (asAntiquities(iIndex)(2) * 30.4) + asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & "0"
										asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
										Else
											sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
										End If
									End If
									bDone = True
								Next
							End If
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If
							aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
							sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5) + Int(aiYears(2) * 365) + (aiMonths(2) * 30.4) + aiDays(2)), "<LL />", aiDays(3))
							If aiDays(5) > 0 Then
								aiYears(5) = Int(aiDays(5) / 365)
								aiDays(5) = aiDays(5) Mod 365
								aiMonths(5) = Int(aiDays(5) / 30.4)
								aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								'Call GetAntiquityFromDays(aiDays(5), lEmployeeStartDate, aiYears(5), aiMonths(5), aiDays(5))
							End If
							aiYears(5) = aiYears(5) + aiYears(2)
							aiMonths(5) = aiMonths(5) + aiMonths(2)
							aiDays(5) = aiDays(5) + aiDays(2)
							If aiDays(5) >= 30 Then
								aiMonths(5) = aiMonths(5) + Int(aiDays(5) / 30)
								aiDays(5) = aiDays(5) Mod 30
							End If
							If aiMonths(5) >= 12 Then
								aiYears(5) = aiYears(5) + Int(aiMonths(5) / 12)
								aiMonths(5) = aiMonths(5) Mod 12
							End If
							Response.Write Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))

							If B_COMMENTS And ((aiYears(5) + aiMonths(5) + aiDays(5)) > 0) Then
								sResultsByIDs = sResultsByIDs & lEmployeeID & "," & aiYears(5) & "," & aiMonths(5) & "," & aiDays(5) & ";"
							End If

							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							sRowContents = "<SPAN COLS=""9"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							iCounter = 0
							sRowContents = "<SPAN COLS=""9"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							iTotalCounter = 0
						End If
						Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""8"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lZoneID = CLng(oRecordset.Fields("ZoneID").Value)
						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""8"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lAreaID <> CLng(oRecordset.Fields("AreaID").Value) Then
						If lAreaID <> -2 Then
							If Not bDone Then
								For iIndex = 0 To UBound(asAntiquities)
									If (asAntiquities(iIndex)(0) = lEmployeeID) Then
										sRowContents = TABLE_SEPARATOR
										sRowContents = sRowContents & "<CENTER>---</CENTER>"
										sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Actividad en " & CleanStringForHTML(asAntiquities(iIndex)(4))
										aiYears(2) = aiYears(2) + asAntiquities(iIndex)(1)
										aiMonths(2) = aiMonths(2) + asAntiquities(iIndex)(2)
										aiDays(2) = aiDays(2) + asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(1)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(2)
										sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & Int(asAntiquities(iIndex)(1) * 365) + (asAntiquities(iIndex)(2) * 30.4) + asAntiquities(iIndex)(3)
										sRowContents = sRowContents & TABLE_SEPARATOR & "0"
										asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
										Else
											sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
										End If
									End If
									bDone = True
								Next
							End If
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If
							aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
							sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5) + Int(aiYears(2) * 365) + (aiMonths(2) * 30.4) + aiDays(2)), "<LL />", aiDays(3))
							If aiDays(5) > 0 Then
								aiYears(5) = Int(aiDays(5) / 365)
								aiDays(5) = aiDays(5) Mod 365
								aiMonths(5) = Int(aiDays(5) / 30.4)
								aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								'Call GetAntiquityFromDays(aiDays(5), lEmployeeStartDate, aiYears(5), aiMonths(5), aiDays(5))
							End If
							aiYears(5) = aiYears(5) + aiYears(2)
							aiMonths(5) = aiMonths(5) + aiMonths(2)
							aiDays(5) = aiDays(5) + aiDays(2)
							If aiDays(5) >= 30 Then
								aiMonths(5) = aiMonths(5) + Int(aiDays(5) / 30)
								aiDays(5) = aiDays(5) Mod 30
							End If
							If aiMonths(5) >= 12 Then
								aiYears(5) = aiYears(5) + Int(aiMonths(5) / 12)
								aiMonths(5) = aiMonths(5) Mod 12
							End If
							Response.Write Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))

							If B_COMMENTS And ((aiYears(5) + aiMonths(5) + aiDays(5)) > 0) Then
								sResultsByIDs = sResultsByIDs & lEmployeeID & "," & aiYears(5) & "," & aiMonths(5) & "," & aiDays(5) & ";"
							End If

							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							sRowContents = "<SPAN COLS=""9"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							iCounter = 0
						End If
						Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""8"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lEmployeeID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If Not bDone Then
							For iIndex = 0 To UBound(asAntiquities)
								If (asAntiquities(iIndex)(0) = lEmployeeID) Then
									sRowContents = TABLE_SEPARATOR
									sRowContents = sRowContents & "<CENTER>---</CENTER>"
									sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Actividad en " & CleanStringForHTML(asAntiquities(iIndex)(4))
									aiYears(2) = aiYears(2) + asAntiquities(iIndex)(1)
									aiMonths(2) = aiMonths(2) + asAntiquities(iIndex)(2)
									aiDays(2) = aiDays(2) + asAntiquities(iIndex)(3)
									sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(1)
									sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(2)
									sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(3)
									sRowContents = sRowContents & TABLE_SEPARATOR & Int(asAntiquities(iIndex)(1) * 365) + (asAntiquities(iIndex)(2) * 30.4) + asAntiquities(iIndex)(3)
									sRowContents = sRowContents & TABLE_SEPARATOR & "0"
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
									Else
										sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
									End If
								End If
								bDone = True
							Next
						End If
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If
						aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
						sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5) + Int(aiYears(2) * 365) + (aiMonths(2) * 30.4) + aiDays(2)), "<LL />", aiDays(3))
						If aiDays(5) > 0 Then
							aiYears(5) = Int(aiDays(5) / 365)
							aiDays(5) = aiDays(5) Mod 365
							aiMonths(5) = Int(aiDays(5) / 30.4)
							aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
							'Call GetAntiquityFromDays(aiDays(5), lEmployeeStartDate, aiYears(5), aiMonths(5), aiDays(5))
						End If
						aiYears(5) = aiYears(5) + aiYears(2)
						aiMonths(5) = aiMonths(5) + aiMonths(2)
						aiDays(5) = aiDays(5) + aiDays(2)
						If aiDays(5) >= 30 Then
							aiMonths(5) = aiMonths(5) + Int(aiDays(5) / 30)
							aiDays(5) = aiDays(5) Mod 30
						End If
						If aiMonths(5) >= 12 Then
							aiYears(5) = aiYears(5) + Int(aiMonths(5) / 12)
							aiMonths(5) = aiMonths(5) Mod 12
						End If
						Response.Write Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))

						If B_COMMENTS And ((aiYears(5) + aiMonths(5) + aiDays(5)) > 0) Then
							sResultsByIDs = sResultsByIDs & lEmployeeID & "," & aiYears(5) & "," & aiMonths(5) & "," & aiDays(5) & ";"
						End If

						sFileContents = ""
						For iIndex = 0 To UBound(aiYears)
							aiYears(iIndex) = 0
							aiMonths(iIndex) = 0
							aiDays(iIndex) = 0
						Next
						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</B>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B>"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</B>")
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B><YY /></B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B><MM /></B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B><DD /></B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B><TT /></B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B><LL /></B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
						Else
							sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
						End If
						lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
						iCounter = iCounter + 1
						iTotalCounter = iTotalCounter + 1
						bDone = False

						sErrorDescription = "No se pudo obtener la información de los registros."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, OcurredDate, EndDate From EmployeesAbsencesLKP Where (AbsenceID In (10,95)) And (EmployeeID=" & lEmployeeID & ") And (EmployeesAbsencesLKP.OcurredDate<=EmployeesAbsencesLKP.EndDate) Order By OcurredDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAbsencesRecordset)
						Do While Not oAbsencesRecordset.EOF
							If CLng(oAbsencesRecordset.Fields("EndDate").Value) > lDate Then
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lDate, -1, -1, -1)
								lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(lDate))) + 1
								aiDays(4) = aiDays(4) + lTemp
								Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), lDate, aiYears(0), aiMonths(0), aiDays(0))
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value), -1, -1, -1)
								lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value)))) + 1
								aiDays(4) = aiDays(4) + lTemp
								Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), CLng(oAbsencesRecordset.Fields("EndDate").Value), aiYears(0), aiMonths(0), aiDays(0))
							End If
							oAbsencesRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oAbsencesRecordset.Close
					End If

					sRowContents = TABLE_SEPARATOR
					If CLng(oRecordset.Fields("EndDate").Value) > lDate Then
						sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(lDate)
						lEndDate = lDate
					Else
						sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
						lEndDate = CLng(oRecordset.Fields("EndDate").Value)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
					aiDays(1) = aiDays(1) + lTemp
					'Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("EmployeeDate").Value), lEndDate, aiYears(0), aiMonths(0), aiDays(0))
					aiDays(0) = lTemp
					aiYears(0) = Int(aiDays(0) / 365)
					aiDays(0) = aiDays(0) Mod 365
					aiMonths(0) = Int(aiDays(0) / 30.4)
					aiDays(0) = Int(aiDays(0) - (aiMonths(0) * 30.4))
					sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
					If CInt(oRecordset.Fields("ActiveEmployeeID").Value) = 0 Then
						aiDays(3) = aiDays(3) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						sRowContents = sRowContents & TABLE_SEPARATOR & Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "0"
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
					Else
						sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
					End If

					lEmployeeStartDate = CLng(oRecordset.Fields("EmployeeDate").Value)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				For iIndex = 0 To UBound(asAntiquities)
					If (asAntiquities(iIndex)(0) = lEmployeeID) Then
						sRowContents = TABLE_SEPARATOR
						sRowContents = sRowContents & "<CENTER>---</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Actividad en " & CleanStringForHTML(asAntiquities(iIndex)(4))
						aiYears(2) = aiYears(2) + asAntiquities(iIndex)(1)
						aiMonths(2) = aiMonths(2) + asAntiquities(iIndex)(2)
						aiDays(2) = aiDays(2) + asAntiquities(iIndex)(3)
						sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(1)
						sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(2)
						sRowContents = sRowContents & TABLE_SEPARATOR & asAntiquities(iIndex)(3)
						sRowContents = sRowContents & TABLE_SEPARATOR & Int(asAntiquities(iIndex)(1) * 365) + (asAntiquities(iIndex)(2) * 30.4) + asAntiquities(iIndex)(3)
						sRowContents = sRowContents & TABLE_SEPARATOR & "0"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
						Else
							sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
						End If
					End If
				Next
				sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
				Else
					sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
				End If
				aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
				sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5) + Int(aiYears(2) * 365) + (aiMonths(2) * 30.4) + aiDays(2)), "<LL />", aiDays(3))
				If aiDays(5) > 0 Then
					aiYears(5) = Int(aiDays(5) / 365)
					aiDays(5) = aiDays(5) Mod 365
					aiMonths(5) = Int(aiDays(5) / 30.4)
					aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
					'Call GetAntiquityFromDays(aiDays(5), lEmployeeStartDate, aiYears(5), aiMonths(5), aiDays(5))
				End If
				aiYears(5) = aiYears(5) + aiYears(2)
				aiMonths(5) = aiMonths(5) + aiMonths(2)
				aiDays(5) = aiDays(5) + aiDays(2)
				If aiDays(5) >= 30 Then
					aiMonths(5) = aiMonths(5) + Int(aiDays(5) / 30)
					aiDays(5) = aiDays(5) Mod 30
				End If
				If aiMonths(5) >= 12 Then
					aiYears(5) = aiYears(5) + Int(aiMonths(5) / 12)
					aiMonths(5) = aiMonths(5) Mod 12
				End If
				Response.Write Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))

				If B_COMMENTS And ((aiYears(5) + aiMonths(5) + aiDays(5)) > 0) Then
					sResultsByIDs = sResultsByIDs & lEmployeeID & "," & aiYears(5) & "," & aiMonths(5) & "," & aiDays(5) & ";"
				End If

				Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
				sRowContents = "<SPAN COLS=""9"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = "<SPAN COLS=""9"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	If B_COMMENTS Then
		Response.Write vbNewLine & "<!--" & vbNewLine & Replace(Replace(sResultsByIDs, ",", vbTab), ";", vbNewLine) & vbNewLine & "-->" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1117 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1118(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To get the records from EmployeesHistoryList and
'         EmployeesAbsencesLKP
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1118"
	Dim oRecordset
	Dim oAbsencesRecordset
	Dim sCondition
	Dim aiYears
	Dim aiMonths
	Dim aiDays
	Dim iIndex
	Dim lDate
	Dim lTemp
	Dim lEndDate
	Dim lZoneID
	Dim lAreaID
	Dim lEmployeeID
	Dim iCounter
	Dim iTotalCounter
	Dim sZoneTitle
	Dim sAreaTitle
	Dim bSkip
	Dim bFirst
	Dim sFileContents
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	aiYears = Split("0,0,0,0,0,0", ",")
	aiMonths = Split("0,0,0,0,0,0", ",")
	aiDays = Split("0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(aiYears)
		aiYears(iIndex) = 0
		aiMonths(iIndex) = 0
		aiDays(iIndex) = 0
	Next
	lDate = CLng(oRequest("PayrollYear").Item & oRequest("PayrollMonth").Item & oRequest("PayrollDay").Item)
	bSkip = False
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	sCondition = sCondition & " And (EmployeeDate<=EmployeesHistoryList.EndDate)"
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees."), "PositionTypes.", "Employees."), "JobTypes.", "Jobs."), "Zones.", "Zones3.")
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Zones1.ZoneID, Zones1.ZoneCode, Zones1.ZoneName, Areas.AreaID, AreaShortName, AreaName, Employees.EmployeeID, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.Active, StatusName, StatusEmployees.Active As StatusActive, Reasons.ActiveEmployeeID From Employees, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones As Zones3, Zones As Zones2, Zones As Zones1 Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (Areas.StartDate<=" & lDate & ") And (Areas.EndDate>=" & lDate & ") And (Zones3.StartDate<=" & lDate & ") And (Zones3.EndDate>=" & lDate & ") And (Employees.StartDate<=" & lDate - 100000 & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & " Order By Zones1.ZoneID, Areas.AreaID, Employees.EmployeeID, EmployeeDate Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No. empleado,Nombre del Empleado,RFC,CURP,Años,Meses,Días,Días antigüedad,Días licencia,Premio por antigüedad,Premio Moneda", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

				lZoneID = -2
				lAreaID = -2
				lEmployeeID = -2
				iCounter = 0
				iTotalCounter = 0
				Do While Not oRecordset.EOF
					If lZoneID <> CLng(oRecordset.Fields("ZoneID").Value) Then
						If lZoneID <> -2 Then
							If Not bSkip Then
								sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
								Else
									sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
								End If
								aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
								sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
								If aiDays(5) > 0 Then
									aiYears(5) = Int(aiDays(5) / 365)
									aiDays(5) = aiDays(5) Mod 365
									aiMonths(5) = Int(aiDays(5) / 30.4)
									aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								End If

								sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
								If aiYears(5) >= 10 Then
									If (aiYears(5) Mod 5) = 0 Then
										If aiYears(5) >= 15 Then
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
										Else
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
										End If

										If Len(sZoneTitle) > 0 Then
											Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
											asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sZoneTitle = ""
										End If

										If Len(sAreaTitle) > 0 Then
											asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sAreaTitle = ""
										End If

										Response.Write sFileContents
										iCounter = iCounter + 1
										iTotalCounter = iTotalCounter + 1
									End If
								End If
							End If
							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							If iCounter > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								iCounter = 0
							End If
							If iTotalCounter > 0 Then
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								iTotalCounter = 0
							End If
						End If
						sZoneTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & "</B>"
						lZoneID = CLng(oRecordset.Fields("ZoneID").Value)
						sAreaTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lAreaID <> CLng(oRecordset.Fields("AreaID").Value) Then
						If lAreaID <> -2 Then
							If Not bSkip Then
								sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
								Else
									sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
								End If
								aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
								sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
								If aiDays(5) > 0 Then
									aiYears(5) = Int(aiDays(5) / 365)
									aiDays(5) = aiDays(5) Mod 365
									aiMonths(5) = Int(aiDays(5) / 30.4)
									aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								End If

								sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
								If aiYears(5) >= 10 Then
									If (aiYears(5) Mod 5) = 0 Then
										If aiYears(5) >= 15 Then
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
										Else
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
										End If

										If Len(sZoneTitle) > 0 Then
											Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
											asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sZoneTitle = ""
										End If

										If Len(sAreaTitle) > 0 Then
											asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sAreaTitle = ""
										End If

										Response.Write sFileContents
										iCounter = iCounter + 1
										iTotalCounter = iTotalCounter + 1
									End If
								End If
							End If

							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							If iCounter > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							End If
							iCounter = 0
						End If
						sAreaTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lEmployeeID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If Not bSkip Then
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If
							aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
							sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
							If aiDays(5) > 0 Then
								aiYears(5) = Int(aiDays(5) / 365)
								aiDays(5) = aiDays(5) Mod 365
								aiMonths(5) = Int(aiDays(5) / 30.4)
								aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
							End If

							sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
							If aiYears(5) >= 10 Then
								If (aiYears(5) Mod 5) = 0 Then
									If aiYears(5) >= 15 Then
										sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
									Else
										sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
									End If

									If Len(sZoneTitle) > 0 Then
										Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
										asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
										sZoneTitle = ""
									End If

									If Len(sAreaTitle) > 0 Then
										asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
										sAreaTitle = ""
									End If

									Response.Write sFileContents
									iCounter = iCounter + 1
									iTotalCounter = iTotalCounter + 1
								End If
							End If
						End If

						sFileContents = ""
						For iIndex = 0 To UBound(aiYears)
							aiYears(iIndex) = 0
							aiMonths(iIndex) = 0
							aiDays(iIndex) = 0
						Next

						lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
						bSkip = False
						bFirst = True
						If Not bSkip Then
							If bFirst And ((CLng(oRecordset.Fields("EndDate").Value) < lDate) Or (CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1) Or (CInt(oRecordset.Fields("Active").Value) <> 1) Or (CInt(oRecordset.Fields("StatusActive").Value) <> 1) Or (CInt(oRecordset.Fields("ActiveEmployeeID").Value) <> 1)) Then bSkip = True
							bFirst = False
							sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</B>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B>"
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</B>")
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><YY /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><MM /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><DD /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><TT /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><LL /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><P1 /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><P2 /></B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If

							sErrorDescription = "No se pudo obtener la información de los registros."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, OcurredDate, EndDate From EmployeesAbsencesLKP Where (AbsenceID In (10,95)) And (EmployeeID=" & lEmployeeID & ") And (EmployeesAbsencesLKP.OcurredDate<=EmployeesAbsencesLKP.EndDate) Order By OcurredDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAbsencesRecordset)
							Do While Not oAbsencesRecordset.EOF
								If CLng(oAbsencesRecordset.Fields("EndDate").Value) > lDate Then
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lDate, -1, -1, -1)
									lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(lDate))) + 1
									aiDays(4) = aiDays(4) + lTemp
									Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), lDate, aiYears(0), aiMonths(0), aiDays(0))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value), -1, -1, -1)
									lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value)))) + 1
									aiDays(4) = aiDays(4) + lTemp
									Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), CLng(oAbsencesRecordset.Fields("EndDate").Value), aiYears(0), aiMonths(0), aiDays(0))
								End If
								oAbsencesRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
							oAbsencesRecordset.Close
						End If
					End If

					If Not bSkip Then
						sRowContents = TABLE_SEPARATOR
						If CLng(oRecordset.Fields("EndDate").Value) > lDate Then
							sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(lDate)
							lEndDate = lDate
						Else
							sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
							lEndDate = CLng(oRecordset.Fields("EndDate").Value)
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						aiDays(1) = aiDays(1) + lTemp

						aiDays(0) = lTemp
						aiYears(0) = Int(aiDays(0) / 365)
						aiDays(0) = aiDays(0) Mod 365
						aiMonths(0) = Int(aiDays(0) / 30.4)
						aiDays(0) = Int(aiDays(0) - (aiMonths(0) * 30.4))
						sRowContents = sRowContents & TABLE_SEPARATOR & aiYears(0)
						sRowContents = sRowContents & TABLE_SEPARATOR & aiMonths(0)
						sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(0)
						sRowContents = sRowContents & TABLE_SEPARATOR & Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						If CInt(oRecordset.Fields("ActiveEmployeeID").Value) = 0 Then
							aiDays(3) = aiDays(3) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
							sRowContents = sRowContents & TABLE_SEPARATOR & Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "0"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & TABLE_SEPARATOR
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
						Else
							sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
						End If
					End If

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				If Not bSkip Then
					sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
					Else
						sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
					End If
					aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
					sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
					If aiDays(5) > 0 Then
						aiYears(5) = Int(aiDays(5) / 365)
						aiDays(5) = aiDays(5) Mod 365
						aiMonths(5) = Int(aiDays(5) / 30.4)
						aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
					End If

					sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
					If aiYears(5) >= 10 Then
						If (aiYears(5) Mod 5) = 0 Then
							If aiYears(5) >= 15 Then
								sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
							Else
								sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
							End If

							If Len(sZoneTitle) > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								sZoneTitle = ""
							End If

							If Len(sAreaTitle) > 0 Then
								asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								sAreaTitle = ""
							End If

							Response.Write sFileContents
							iCounter = iCounter + 1
							iTotalCounter = iTotalCounter + 1
						End If
					End If
				End If
				If iCounter > 0 Then
					Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
					sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
				End If
				If iTotalCounter > 0 Then
					sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
				Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	Set oAbsencesRecordset = Nothing
	BuildReport1118 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1119(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To get the records from EmployeesHistoryList and
'         EmployeesAbsencesLKP
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1119"
	Dim oRecordset
	Dim oAbsencesRecordset
	Dim sCondition
	Dim aiYears
	Dim aiMonths
	Dim aiDays
	Dim iIndex
	Dim lPayrollID
	Dim lForPayrollID
	Dim lTemp
	Dim lEndDate
	Dim lZoneID
	Dim lAreaID
	Dim lEmployeeID
	Dim iCounter
	Dim iTotalCounter
	Dim sZoneTitle
	Dim sAreaTitle
	Dim bSkip
	Dim bFirst
	Dim sFileContents
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	aiYears = Split("0,0,0,0,0,0", ",")
	aiMonths = Split("0,0,0,0,0,0", ",")
	aiDays = Split("0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(aiYears)
		aiYears(iIndex) = 0
		aiMonths(iIndex) = 0
		aiDays(iIndex) = 0
	Next
	bSkip = False
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	sCondition = sCondition & " And (EmployeeDate<=EmployeesHistoryList.EndDate)"
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees."), "PositionTypes.", "Employees."), "JobTypes.", "Jobs."), "Zones.", "Zones3.")
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Zones1.ZoneID, Zones1.ZoneCode, Zones1.ZoneName, Areas.AreaID, AreaShortName, AreaName, Employees.EmployeeID, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.Active, StatusName, StatusEmployees.Active As StatusActive, Reasons.ActiveEmployeeID From Employees, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones As Zones3, Zones As Zones2, Zones As Zones1 Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones3.StartDate<=" & lForPayrollID & ") And (Zones3.EndDate>=" & lForPayrollID & ") And (Employees.StartDate<=" & lForPayrollID - 100000 & ") " & sCondition & " Order By Zones1.ZoneID, Areas.AreaID, Employees.EmployeeID, EmployeeDate Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No. empleado,Nombre del Empleado,RFC,CURP,Años,Meses,Días,Días antigüedad,Días licencia,Premio por antigüedad,Premio Moneda", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

				lZoneID = -2
				lAreaID = -2
				lEmployeeID = -2
				iCounter = 0
				iTotalCounter = 0
				Do While Not oRecordset.EOF
					If lZoneID <> CLng(oRecordset.Fields("ZoneID").Value) Then
						If lZoneID <> -2 Then
							If Not bSkip Then
								sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
								Else
									sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
								End If
								aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
								sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
								If aiDays(5) > 0 Then
									aiYears(5) = Int(aiDays(5) / 365)
									aiDays(5) = aiDays(5) Mod 365
									aiMonths(5) = Int(aiDays(5) / 30.4)
									aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								End If

								sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
								If aiYears(5) >= 10 Then
									If (aiYears(5) Mod 5) = 0 Then
										If aiYears(5) >= 15 Then
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
										Else
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
										End If

										If Len(sZoneTitle) > 0 Then
											Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
											asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sZoneTitle = ""
										End If

										If Len(sAreaTitle) > 0 Then
											asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sAreaTitle = ""
										End If

										Response.Write sFileContents
										iCounter = iCounter + 1
										iTotalCounter = iTotalCounter + 1
									End If
								End If
							End If
							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							If iCounter > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								iCounter = 0
							End If
							If iTotalCounter > 0 Then
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								iTotalCounter = 0
							End If
						End If
						sZoneTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & "</B>"
						lZoneID = CLng(oRecordset.Fields("ZoneID").Value)
						sAreaTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lAreaID <> CLng(oRecordset.Fields("AreaID").Value) Then
						If lAreaID <> -2 Then
							If Not bSkip Then
								sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
								Else
									sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
								End If
								aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
								sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
								If aiDays(5) > 0 Then
									aiYears(5) = Int(aiDays(5) / 365)
									aiDays(5) = aiDays(5) Mod 365
									aiMonths(5) = Int(aiDays(5) / 30.4)
									aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
								End If

								sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
								If aiYears(5) >= 10 Then
									If (aiYears(5) Mod 5) = 0 Then
										If aiYears(5) >= 15 Then
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
										Else
											sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
										End If

										If Len(sZoneTitle) > 0 Then
											Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
											asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sZoneTitle = ""
										End If

										If Len(sAreaTitle) > 0 Then
											asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
											If bForExport Then
												lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
											Else
												lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
											End If
											sAreaTitle = ""
										End If

										Response.Write sFileContents
										iCounter = iCounter + 1
										iTotalCounter = iTotalCounter + 1
									End If
								End If
							End If

							sFileContents = ""
							For iIndex = 0 To UBound(aiYears)
								aiYears(iIndex) = 0
								aiMonths(iIndex) = 0
								aiDays(iIndex) = 0
							Next
							If iCounter > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
								asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							End If
							iCounter = 0
						End If
						sAreaTitle = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & "</B>" & TABLE_SEPARATOR & "<SPAN COLS=""10"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</B>"
						lAreaID = CLng(oRecordset.Fields("AreaID").Value)
					End If
					If lEmployeeID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If Not bSkip Then
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<SPAN COLS=""2"" />Ausencias" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR & aiDays(4) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If
							aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
							sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
							If aiDays(5) > 0 Then
								aiYears(5) = Int(aiDays(5) / 365)
								aiDays(5) = aiDays(5) Mod 365
								aiMonths(5) = Int(aiDays(5) / 30.4)
								aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
							End If

							sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
							If aiYears(5) >= 10 Then
								If (aiYears(5) Mod 5) = 0 Then
									If aiYears(5) >= 15 Then
										sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
									Else
										sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
									End If

									If Len(sZoneTitle) > 0 Then
										Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
										asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
										sZoneTitle = ""
									End If

									If Len(sAreaTitle) > 0 Then
										asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
										sAreaTitle = ""
									End If

									Response.Write sFileContents
									iCounter = iCounter + 1
									iTotalCounter = iTotalCounter + 1
								End If
							End If
						End If

						sFileContents = ""
						For iIndex = 0 To UBound(aiYears)
							aiYears(iIndex) = 0
							aiMonths(iIndex) = 0
							aiDays(iIndex) = 0
						Next

						lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
						bSkip = False
						bFirst = True
						If Not bSkip Then
							If bFirst And ((CLng(oRecordset.Fields("EndDate").Value) < lForPayrollID) Or (CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1) Or (CInt(oRecordset.Fields("Active").Value) <> 1) Or (CInt(oRecordset.Fields("StatusActive").Value) <> 1) Or (CInt(oRecordset.Fields("ActiveEmployeeID").Value) <> 1)) Then bSkip = True
							bFirst = False
							sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</B>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B>"
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</B>")
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><YY /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><MM /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><DD /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><TT /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><LL /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><P1 /></B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B><P2 /></B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								sFileContents = sFileContents & GetTableRowText(asRowContents, True, sErrorDescription)
							Else
								sFileContents = sFileContents & GetTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", sErrorDescription)
							End If

							sErrorDescription = "No se pudo obtener la información de los registros."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, OcurredDate, EndDate From EmployeesAbsencesLKP Where (AbsenceID In (10,95)) And (EmployeeID=" & lEmployeeID & ") Order By OcurredDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAbsencesRecordset)
							Do While Not oAbsencesRecordset.EOF
								If CLng(oAbsencesRecordset.Fields("EndDate").Value) > lForPayrollID Then
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1)
									lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(lForPayrollID))) + 1
									aiDays(4) = aiDays(4) + lTemp
									Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), lForPayrollID, aiYears(0), aiMonths(0), aiDays(0))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value), -1, -1, -1)
									lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oAbsencesRecordset.Fields("EndDate").Value)))) + 1
									aiDays(4) = aiDays(4) + lTemp
									Call GetAntiquityFromSerialDates(CLng(oAbsencesRecordset.Fields("OcurredDate").Value), CLng(oAbsencesRecordset.Fields("EndDate").Value), aiYears(0), aiMonths(0), aiDays(0))
								End If
								oAbsencesRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
							oAbsencesRecordset.Close
						End If
					End If

					If Not bSkip Then
						If CLng(oRecordset.Fields("EndDate").Value) > lForPayrollID Then
							lEndDate = lForPayrollID
						Else
							lEndDate = CLng(oRecordset.Fields("EndDate").Value)
						End If
						lTemp = Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						aiDays(1) = aiDays(1) + lTemp
						aiDays(0) = lTemp
						aiYears(0) = Int(aiDays(0) / 365)
						aiDays(0) = aiDays(0) Mod 365
						aiMonths(0) = Int(aiDays(0) / 30.4)
						aiDays(0) = Int(aiDays(0) - (aiMonths(0) * 30.4))
						If CInt(oRecordset.Fields("ActiveEmployeeID").Value) = 0 Then
							aiDays(3) = aiDays(3) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
						End If
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				If Not bSkip Then
					aiDays(5) = aiDays(1) - aiDays(3) - aiDays(4)
					sFileContents = Replace(Replace(sFileContents, "<TT />", aiDays(5)), "<LL />", aiDays(3) + aiDays(4))
					If aiDays(5) > 0 Then
						aiYears(5) = Int(aiDays(5) / 365)
						aiDays(5) = aiDays(5) Mod 365
						aiMonths(5) = Int(aiDays(5) / 30.4)
						aiDays(5) = Int(aiDays(5) - (aiMonths(5) * 30.4))
					End If

					sFileContents = Replace(Replace(Replace(sFileContents, "<YY />", aiYears(5)), "<MM />", aiMonths(5)), "<DD />", aiDays(5))
					If aiYears(5) >= 10 Then
						If (aiYears(5) Mod 5) = 0 Then
							If aiYears(5) >= 15 Then
								sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", Int(aiYears(5) / 5) * 5)
							Else
								sFileContents = Replace(Replace(sFileContents, "<P1 />", Int(aiYears(5) / 5) * 5), "<P2 />", "<CENTER>---</CENTER>")
							End If

							If Len(sZoneTitle) > 0 Then
								Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
								asRowContents = Split(sZoneTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								sZoneTitle = ""
							End If

							If Len(sAreaTitle) > 0 Then
								asRowContents = Split(sAreaTitle, TABLE_SEPARATOR, -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
								sAreaTitle = ""
							End If

							Response.Write sFileContents
							iCounter = iCounter + 1
							iTotalCounter = iTotalCounter + 1
						End If
					End If
				End If
				If iCounter > 0 Then
					Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
					sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR CENTRO DE TRABAJO: " & iCounter & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
				End If
				If iTotalCounter > 0 Then
					sRowContents = "<SPAN COLS=""11"" /><B>TOTAL POR ENTIDAD FEDARATIVA: " & iTotalCounter & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
				Call DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	Set oAbsencesRecordset = Nothing
	BuildReport1119 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1120(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the monthly payroll resume based on CLCs
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1403"
	
	Dim oRecordset
	Dim sRowContents
	Dim lErrorNumber
	Dim sQuery
	Dim lPeriod
	Dim lPayrollID
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
    Dim sYear
    Dim sYearTable
    Dim sExtraConditions
    Dim sExtraTables
    Dim lTotConcept
    Dim lTotEne
    Dim lTotFeb
    Dim lTotMar 
    Dim lTotAbr
    Dim lTotMay 
    Dim lTotJun
    Dim lTotJul
    Dim lTotAgo
    Dim lTotSep
    Dim lTotOct
    Dim lTotNov
    Dim lTotDic
    Dim lTotal
    Dim IsDeduction
    Dim lConcept
    Dim i
    Dim lTypeMax
    Dim sConceptName
    Dim sTypeConcentrate
    Dim sCompanyName
    Dim sTitleReport
    
    Dim lTotalLiq() 
    Redim lTotalLiq(13)
    
	sYear = oRequest("YearID").Item
    lConcept = oRequest("ConceptID").Item

    'If StrComp(lConcept,"-1", vbBinaryCompare) = 0 Then
    '    IsDeduction = 0
    'Else
    '    IsDeduction = 1
    'End If

    'sQuery = "Select ConceptName from Concepts Where ConceptID = " & lConcept
    'lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
    'lConcept = CStr(oRecordset.Fields("ConceptName").Value)
	'sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
    
    sRowContents = ""
    sExtraConditions = ""
    sExtraTables=""
    lTotConcept = 0
    lTotEne=0
    lTotFeb=0
    lTotMar=0
    lTotAbr=0
    lTotMay=0
    lTotJun=0
    lTotJul=0
    lTotAgo=0
    lTotSep=0
    lTotOct=0
    lTotNov=0
    lTotDic=0
    lTotal=0
    IsDeduction = 0

    If Len(lConcept)=0 Then
        lTypeMax = 2
        sConceptName = "Percepciones"
    Else
        lTypeMax = 1
    End If

    For i=0 To ubound(lTotalLiq)
        lTotalLiq(i)=0
    Next
    
    If StrComp(oRequest("ConcentrateConceptID").Item,"0", vbBinaryCompare) = 0 Then
        sYearTable = sYear&"0"
        sTypeConcentrate="CANCELACION"
        sTitleReport = "CONCENTRADO DE CONCEPTOS CANCELADOS"
    Else
        sTypeConcentrate="CIRCULANTE"
        sTitleReport = "CONCENTRADO DE CONCEPTOS EMITIDOS"
        sYearTable = sYear
    End IF

    If Len(oRequest("CompanyID").Item)>0 Then
        sQuery = "Select CompanyName From Companies where CompanyID =" & oRequest("CompanyID").Item
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        If lErrorNumber = 0 Then
            sCompanyName = oRecordset.Fields("CompanyName").Value
        End If
        sExtraConditions = sExtraConditions & "And p.employeeID = ehl.employeeID And ehl.companyID =" & oRequest("CompanyID").Item
        sExtraTables = sExtraTables & ", employeesHistoryListForPayroll ehl "
    End If

    FOR i = 1 TO lTypeMax
        If Len(lConcept)<>0 And StrComp(lConcept,"-1", vbBinaryCompare) = 0 Then
            IsDeduction = 0
            sConceptName = "Percepciones"
            'sExtraConditions = sExtraConditions & "And ConceptID = -1"
        ElseIf Len(lConcept)<>0 And StrComp(lConcept,"-2", vbBinaryCompare) = 0 Then 
            IsDeduction = 1
            sConceptName = "Deducciones"
            'sExtraConditions = sExtraConditions & "And ConceptID = -2"
        End If

               
        sQuery = "Select P.ConceptID,C.ConceptShortName as CPT, C.ConceptName as CPTName, Sum(Case When RecordDate Between "&sYear&"0101 And "&sYear&"0131 Then ConceptAmount Else 0 End )As Enero, sum(Case When RecordDate Between "&sYear&"0201 And "&sYear&"0231 Then ConceptAmount Else 0 End )as Febrero, sum(Case When RecordDate between "&sYear&"0301 and "&sYear&"0331 Then conceptAmount Else 0 End )as Marzo, sum(Case When RecordDate between "&sYear&"0401 and "&sYear&"0431 Then conceptAmount Else 0 End )as Abril, sum(Case When RecordDate between "&sYear&"0501 and "&sYear&"0531 Then conceptAmount Else 0 End )as Mayo, sum(Case When RecordDate between "&sYear&"0601 and "&sYear&"0631 Then conceptAmount Else 0 End )as Junio, sum(Case When RecordDate between "&sYear&"0701 and "&sYear&"0731 Then conceptAmount Else 0 End )as Julio, sum(Case When RecordDate between "&sYear&"0801 and "&sYear&"0831 Then conceptAmount Else 0 End )as Agosto, sum(Case When RecordDate between "&sYear&"0901 and "&sYear&"0931 Then conceptAmount Else 0 End )as Septiembre, sum(Case When RecordDate between "&sYear&"1001 and "&sYear&"1031 Then conceptAmount Else 0 End )as Octubre, sum(Case When RecordDate between "&sYear&"1101 and "&sYear&"1131 Then conceptAmount Else 0 End )as Noviembre, sum(Case When RecordDate between "&sYear&"1201 and "&sYear&"1231 Then conceptAmount Else 0 End )as Diciembre From payroll_"&sYearTable&" p, concepts c "&sExtraTables&" where C.ConceptID = p.ConceptID And p.ConceptID In (Select ConceptID from concepts where ISDEDUCTION = "&IsDeduction&" And ConceptId>0) "&sExtraConditions&" group by  p.ConceptID,c.ConceptShortName, C.ConceptName,OrderInList Order by p.ConceptID,c.ConceptShortName,OrderInList"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        If lErrorNumber = 0 Then
            If Not oRecordset.EOF Then
               If i=1 Then
                    sDate = GetSerialNumberForDate("")
    			    sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
    			    lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
    			    sFilePath = sFilePath & "\"
    			    sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
    			    Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
    			    Response.Flush()
    			    sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
                   
               End If

                sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
                        sRowContents = sRowContents & "<TR>"
                        sRowContents = sRowContents & "<TD rowspan=4 ><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/Logo_ISSSTE.gif"" WIDTH=""50"" HEIGHT=""60""  BORDER=""0"" /></TD>"
                        sRowContents = sRowContents & "<TD COLSPAN=2 FACE=""verdana"">DIRECCIÓN DE ADMINISTRACIÓN</TD><TD COLSPAN=3><B>"&sCompanyName&"</B></TD></TR>"  
                        sRowContents = sRowContents & "<TR><TD COLSPAN=2 FACE=""verdana"">SUBDIRECCIÓN DE PERSONAL</TD><TD COLSPAN=3>"&sTitleReport&"</TD></TR>" 
                        sRowContents = sRowContents & "<TR><TD COLSPAN=5 FACE=""verdana"">JEFATURA DE SERVICIOS DE PERSONAL</TD></TR>" 
                        sRowContents = sRowContents & "<TR><TD COLSPAN=5 FACE=""verdana"">DEPARTAMENTO DE CONTROL DE CIFRAS</TD></TR>" 
                        'sRowContents = sRowContents & "<TD ></TD></TR>"    
                    sRowContents = sRowContents & "</TABLE>"                    
                    sRowContents = sRowContents & "</br>"

               sRowContents =  sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
			        sRowContents = sRowContents & "<TR>"
                        sRowContents = sRowContents & "<TD><B>CONCEPTO</B></TD>"
		                sRowContents = sRowContents & "<TD valign=""middle""><B>"&sTypeConcentrate&" "&sYear&" ("& sConceptName &")</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'ENERO "&sYear&"</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'FEBRERO "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>'MARZO "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>'ABRIL "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>'MAYO "&sYear&"</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'JUNIO "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>'JULIO "&sYear&"</B></TD>"				
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'AGOSTO "&sYear&"</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'SEPTIEMBRE "&sYear&"</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'OCTUBRE "&sYear&"</B></TD>"
				        sRowContents = sRowContents & "<TD valign=""middle""><B>'NOVIEMBRE "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>'DICIEMBRE "&sYear&"</B></TD>"
                        sRowContents = sRowContents & "<TD valign=""middle""><B>TOTAL</B></TD>"
			        sRowContents = sRowContents & "</TR>"

                    Do While Not oRecordset.EOF
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD>"& CStr(oRecordset.Fields("CPT").Value) &"</TD>"
                            sRowContents = sRowContents & "<TD>"& CStr(oRecordset.Fields("CPTName").Value) &"</TD>"
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Enero").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Enero").Value)
                            lTotEne = lTotEne + CDbl(oRecordset.Fields("Enero").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Febrero").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Febrero").Value)
                            lTotFeb=lTotFeb + CDbl(oRecordset.Fields("Febrero").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Marzo").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Marzo").Value)
                            lTotMar=lTotMar + CDbl(oRecordset.Fields("Marzo").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Abril").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Abril").Value)
                            lTotAbr = lTotAbr + CDbl(oRecordset.Fields("Abril").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Mayo").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Mayo").Value)
                            lTotMay = lTotMay + CDbl(oRecordset.Fields("Mayo").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Junio").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Junio").Value)
                            lTotJun = lTotJun + CDbl(oRecordset.Fields("Junio").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Julio").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Julio").Value)
                            lTotJul = lTotJul + CDbl(oRecordset.Fields("Julio").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Agosto").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Agosto").Value)
                            lTotAgo = lTotAgo + CDbl(oRecordset.Fields("Agosto").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Septiembre").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Septiembre").Value)
                            lTotSep = lTotSep + CDbl(oRecordset.Fields("Septiembre").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Octubre").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Octubre").Value)
                            lTotOct = lTotOct + CDbl(oRecordset.Fields("Octubre").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Noviembre").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Noviembre").Value)
                            lTotNov = lTotNov + CDbl(oRecordset.Fields("Noviembre").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Diciembre").Value),2,True,True,True) & "</TD>"
                            lTotConcept = lTotConcept+ CDbl(oRecordset.Fields("Diciembre").Value)
                            lTotDic = lTotDic + CDbl(oRecordset.Fields("Diciembre").Value)
                            sRowContents = sRowContents & "<TD>" & FormatNumber(lTotConcept,2,True,True,True) & "</TD>"
                        sRowContents = sRowContents & "</TR>"
                    oRecordset.MoveNext
                    lTotal =lTotal+lTotConcept
                    lTotConcept = 0
                    Loop

                    sRowContents = sRowContents & "<TR>"
                        sRowContents = sRowContents & "<TD></TD>"
                        sRowContents = sRowContents & "<TD><B>TOTAL " & UCase(sConceptName) & " ORGANOS ISSSTE</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotEne,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotFeb,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotMar,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotAbr,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotMay,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotJun,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotJul,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotAgo,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotSep,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotOct,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotNov,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotDic,2,True,True,True)&"</B></TD>"
                        sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotal,2,True,True,True)&"</B></TD>"
                    sRowContents = sRowContents & "</TR>"
                    If Len(lConcept)=0  Then
                        If i=1  Then 
                                lTotalLiq(0) = lTotalLiq(0) - lTotEne
                                lTotalLiq(1) = lTotalLiq(1) - lTotFeb
                                lTotalLiq(2) = lTotalLiq(2) - lTotMar
                                lTotalLiq(3) = lTotalLiq(3) - lTotAbr
                                lTotalLiq(4) = lTotalLiq(4) - lTotMay
                                lTotalLiq(5) = lTotalLiq(5) - lTotJun
                                lTotalLiq(6) = lTotalLiq(6) - lTotJul
                                lTotalLiq(7) = lTotalLiq(7) - lTotAgo
                                lTotalLiq(8) = lTotalLiq(8) - lTotSep
                                lTotalLiq(9) = lTotalLiq(9) - lTotOct
                                lTotalLiq(10) = lTotalLiq(10) - lTotNov
                                lTotalLiq(11) = lTotalLiq(11) - lTotDic
                                lTotalLiq(12) = lTotalLiq(12) - lTotal

                        Else 
                                lTotalLiq(0) = (lTotalLiq(0) + lTotEne)*-1  
                                lTotalLiq(1) = (lTotalLiq(1) + lTotFeb)*-1  
                                lTotalLiq(2) = (lTotalLiq(2) + lTotMar)*-1  
                                lTotalLiq(3) = (lTotalLiq(3) + lTotAbr)*-1  
                                lTotalLiq(4) = (lTotalLiq(4) + lTotMay)*-1  
                                lTotalLiq(5) = (lTotalLiq(5) + lTotJun)*-1  
                                lTotalLiq(6) = (lTotalLiq(6) + lTotJul)*-1  
                                lTotalLiq(7) = (lTotalLiq(7) + lTotAgo)*-1  
                                lTotalLiq(8) = (lTotalLiq(8) + lTotSep)*-1  
                                lTotalLiq(9) = (lTotalLiq(9) + lTotOct)*-1  
                                lTotalLiq(10) = (lTotalLiq(10) + lTotNov)*-1  
                                lTotalLiq(11) = (lTotalLiq(11) + lTotDic)*-1  
                                lTotalLiq(12) = (lTotalLiq(12) + lTotal)*-1  
                        End IF
                    End if

                sRowContents = sRowContents & "</TABLE>"
                sRowContents = sRowContents & "<br/><br/><br/>"
            End If
        Else
            lErrorNumber = -1
		    sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."    
        End If
        IsDeduction = 1   
        sConceptName = "Deducciones"
        lTotEne=0
        lTotFeb=0
        lTotMar=0
        lTotAbr=0
        lTotMay=0
        lTotJun=0
        lTotJul=0
        lTotAgo=0
        lTotSep=0
        lTotOct=0
        lTotNov=0
        lTotDic=0
        lTotal=0    
    NEXT 
   

    If ((lErrorNumber = 0) And Len(lConcept)=0 ) Then
        sRowContents = sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
            sRowContents = sRowContents & "<TR>"
                sRowContents = sRowContents & "<TD></TD>"
                sRowContents = sRowContents & "<TD><B>TOTAL LIQUIDO DEL ISSSTE "& sYear &"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(0),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(1),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(2),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(3),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(4),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(5),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(6),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(7),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(8),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(9),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(10),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(11),2,True,True,True)&"</B></TD>"
                sRowContents = sRowContents & "<TD><B>"&FormatNumber(lTotalLiq(12),2,True,True,True)&"</B></TD>"
            sRowContents = sRowContents & "</TR>"
        sRowContents = sRowContents & "</TABLE>"
    End If
    
    'Fin Archivo
   
    lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
    If lErrorNumber = 0 Then
            lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
            lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
            oEndDate = Now()
            If (lErrorNumber = 0) And B_USE_SMTP Then
		        If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		    End If    
    Else
        lErrorNumber = -1
		sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
    End If

	Set oRecordset = Nothing
	BuildReport1120 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1126(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de incidencias registradas para el personal filtrado por
'         número de empleado, áreas y período específico
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1126"
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
	'Dim asStateNames
	Dim asAbsenceNames
	Dim asPath
	Dim iCount
	Dim aiAbscenceTotals
	'Dim aiAbscenceGrandTotals
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

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(AbsenceID) As Max From Absences", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value)
		End If
	End If
	For iMin = 0 To iMax
		asAbsenceNames = asAbsenceNames & LIST_SEPARATOR & ""
		aiAbscenceTotals = aiAbscenceTotals & LIST_SEPARATOR & "0"
	Next
	asAbsenceNames = Split(asAbsenceNames, LIST_SEPARATOR)
	aiAbscenceTotals = Split(aiAbscenceTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiAbscenceTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, AbsenceShortName From Absences Where (AbsenceID>-1) And (AbsenceID<100) Order By AbsenceID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition & sCondition2 & " Order By Employees.EmployeeID", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & sCondition & sCondition2 & " Order By Employees.EmployeeID" & " -->" & vbNewLine
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

				iCount = 0
				lTotalForReport = 0
				Do While Not oRecordset.EOF
					iCount = iCount + 1
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
						'aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("JustificationID").Value)) + 1
					Else
						aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
						'aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) = aiAbscenceGrandTotals(CInt(oRecordset.Fields("AbsenceID").Value)) + 1
					End If
					bFirst = True
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If (bFirst) And (lCurrentPaymentCenterID <> CLng(oRecordset.Fields("PaymentCenterID").Value)) Then
					sRowContents = "</TABLE>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					sRowContents = "<BR /><B>TOTALES DEL REPORTE</B><BR />"
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
							'lTotalForArea = lTotalForArea + lTotal
							lTotalForReport = lTotalForReport + lTotal
							sAbsenceShortName = Trim(asAbsenceNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sAbsenceShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
				End If
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				sRowContents = "<BR /><B>REGISTROS TOTALES DEL REPORTE: " & lTotalForReport & "</B><BR />"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.Close
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	BuildReport1126 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1151(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paid amounts for every employee in
'         the given year
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1151"
	Dim sPayrollTitles
	Dim lMaxConceptID
	Dim sConceptTitles
	Dim sConceptAmounts
	Dim sAmounts
	Dim sCurrentAmounts
	Dim adTotalAmounts
	Dim dPerceptionsAmount
	Dim dDeductionsAmount
	Dim sCondition
	Dim oRecordset
	Dim lCurrentID
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim sFontBegin
	Dim sFontEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	oStartDate = Now()
	lMaxConceptID = 0
	sCondition = ""
	If Len(oRequest("ConceptID").Item) > 0 Then sCondition = " And (ConceptID In (" & Replace(oRequest("ConceptID").Item, " ", "") & "))"
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptID) From Concepts Where (EndDate=30000000)" & sCondition, "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lMaxConceptID = CLng(oRecordset.Fields(0).Value)
		End If
		oRecordset.Close
	End If

	sConceptTitles = ""
	sConceptAmounts = ""
	adTotalAmounts = ""
	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptShortName, ConceptName From Concepts Where (EndDate=30000000) " & sCondition & " Order By IsDeduction, OrderInList, ConceptShortName", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sConceptTitles = sConceptTitles & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & "<BR />"
			sConceptAmounts = sConceptAmounts & "<CONCEPT_AMOUNT_" & CStr(oRecordset.Fields("ConceptID").Value) & "_YYYYMMDD /><BR />"
			adTotalAmounts = adTotalAmounts & CStr(oRecordset.Fields("ConceptID").Value) & ",0,0;"
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		If Len(adTotalAmounts) > 0 Then adTotalAmounts = Left(adTotalAmounts, (Len(adTotalAmounts) - Len(";")))
		oRecordset.Close
	End If
	adTotalAmounts = Split(adTotalAmounts, ";")
	For iIndex = 0 To UBound(adTotalAmounts)
		adTotalAmounts(iIndex) = Split(adTotalAmounts(iIndex), ",")
		adTotalAmounts(iIndex)(0) = CLng(adTotalAmounts(iIndex)(0))
	Next

	sPayrollTitles = ""
	sAmounts = ""
	sErrorDescription = "No se pudieron obtener los registros de las nóminas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct RecordDate From Payroll_" & oRequest("YearID").Item & " Where (RecordDate>" & oRequest("YearID").Item & "0000) Order By RecordDate", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sPayrollTitles = sPayrollTitles & "<TD ALIGN=""CENTER""><B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("RecordDate").Value), -1, -1, -1) & "</B></TD>"
			sAmounts = sAmounts & TABLE_SEPARATOR & Replace(sConceptAmounts, "YYYYMMDD", CStr(oRecordset.Fields("RecordDate").Value))
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		sAmounts = sAmounts & TABLE_SEPARATOR & Replace(sConceptAmounts, "YYYYMMDD", "Count") & TABLE_SEPARATOR & Replace(sConceptAmounts, "YYYYMMDD", "Total")
		oRecordset.Close
	End If

	If lErrorNumber = 0 Then
		sCondition = ""
		Call GetConditionFromURL(oRequest, sCondition, -1, -1)
		sCondition = Replace(sCondition, "(Companies.", "(Employees.")
		sCondition = Replace(sCondition, "(EmployeeTypes.", "(Employees.")
		sCondition = Replace(sCondition, "(PositionTypes.", "(Employees.")
		sCondition = Replace(sCondition, "(Journeys.", "(Employees.")
		sCondition = Replace(sCondition, "(Shifts.", "(Employees.")
		sCondition = Replace(sCondition, "(Levels.", "(Employees.")
		sCondition = Replace(sCondition, "(PaymentCenters.", "(Employees.")
		sCondition = Replace(sCondition, "(Positions.", "(Jobs.")
		sErrorDescription = "No se pudieron obtener los acumulados anuales de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Payroll_" & oRequest("YearID").Item & ".ConceptAmount, Payroll_" & oRequest("YearID").Item & ".RecordDate, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Areas.EconomiczoneID, PositionShortName From Payroll_" & oRequest("YearID").Item & ", Concepts, Employees, Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Payroll_" & oRequest("YearID").Item & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And (Concepts.EndDate=30000000) And (Services.EndDate=30000000) And (EmployeeTypes.EndDate=30000000) And (PositionTypes.EndDate=30000000) And (GroupGradeLevels.EndDate=30000000) And (Journeys.EndDate=30000000) And (Shifts.EndDate=30000000) And (Levels.EndDate=30000000) And (PaymentCenters.EndDate=30000000) And (Jobs.EndDate=30000000) And (Zones.EndDate=30000000) And (Zones01.EndDate=30000000) And (Zones02.EndDate=30000000) And (Areas.EndDate=30000000) And (GeneratingAreas.EndDate=30000000) And (Positions.EndDate=30000000) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (RecordDate>" & oRequest("YearID").Item & "0000) " & sCondition & " Order By Employees.EmployeeID, Payroll_" & oRequest("YearID").Item & ".RecordDate, Concepts.IsDeduction, Concepts.OrderInList, Concepts.ConceptShortName", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sDate = GetSerialNumberForDate("")
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN)
				If Not FolderExists(Server.MapPath(sFileName), sErrorDescription) Then lErrorNumber = CreateFolder(Server.MapPath(sFileName), sErrorDescription)
				sFileName = sFileName & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
				If lErrorNumber = 0 Then
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & ".zip" & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()
					sFileName = Server.MapPath(sFileName)

					lErrorNumber = AppendTextToFile(sFileName & ".xls", "<TABLE BORDER=""1"">", sErrorDescription)
					sRowContents = "<TR><TD ALIGN=""CENTER""><B>No. del empleado</B></TD><TD ALIGN=""CENTER""><B>Apellido paterno</B></TD><TD ALIGN=""CENTER""><B>Apellido materno</B></TD><TD ALIGN=""CENTER""><B>Nombre</B></TD><TD ALIGN=""CENTER""><B>RFC</B></TD><TD ALIGN=""CENTER""><B>Zona Geográfica</B></TD><TD ALIGN=""CENTER""><B>Puesto</B></TD><TD ALIGN=""CENTER""><B>Conceptos</B></TD>" & sPayrollTitles & "<TD ALIGN=""CENTER""><B>No. de pagos</B></TD><TD ALIGN=""CENTER""><B>TOTAL</B></TD></TR>"
					lErrorNumber = AppendTextToFile(sFileName & ".xls", sRowContents, sErrorDescription)
						lCurrentID = -2
						dPerceptionsAmount = 0
						dDeductionsAmount = 0
						Do While Not oRecordset.EOF
							If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								If lCurrentID > -1 Then
									For iIndex = 0 To UBound(adTotalAmounts)
										sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & adTotalAmounts(iIndex)(0) & "_Count />", "<B>" & adTotalAmounts(iIndex)(1) & "</B>")
										sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & adTotalAmounts(iIndex)(0) & "_Total />", "<B>" & FormatNumber(adTotalAmounts(iIndex)(2), 2, True, False, True) & "</B>")
									Next
									asRowContents = Split(sRowContents & sCurrentAmounts, TABLE_SEPARATOR, -1, vbBinaryCompare)
									lErrorNumber = AppendTextToFile(sFileName & ".xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
								End If
								lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								For iIndex = 0 To UBound(adTotalAmounts)
									adTotalAmounts(iIndex)(1) = 0
									adTotalAmounts(iIndex)(2) = 0
								Next
								dPerceptionsAmount = 0
								dDeductionsAmount = 0
								sCurrentAmounts = sAmounts
								sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & " "
								End If
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
								sRowContents = sRowContents & TABLE_SEPARATOR & sConceptTitles
							End If
							sFontBegin = ""
							sFontEnd = ""
							If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
'''							sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_NAME_" & CStr(oRecordset.Fields("RecordDate").Value) & " />", sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & sFontEnd & "<BR /><CONCEPT_NAME_" & CStr(oRecordset.Fields("RecordDate").Value) & " />")
							sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & CStr(oRecordset.Fields("ConceptID").Value) & "_" & CStr(oRecordset.Fields("RecordDate").Value) & " />", sFontBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sFontEnd)
							For iIndex = 0 To UBound(adTotalAmounts)
								If adTotalAmounts(iIndex)(0) = CLng(oRecordset.Fields("ConceptID").Value) Then
									adTotalAmounts(iIndex)(1) = adTotalAmounts(iIndex)(1) + 1
									adTotalAmounts(iIndex)(2) = adTotalAmounts(iIndex)(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
									Exit For
								End If
							Next
							'If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
							'	dPerceptionsAmount = dPerceptionsAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							'ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
							'	dDeductionsAmount = dDeductionsAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
							'End If
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						If False Then
							For iIndex = 0 To lMaxConceptID
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0115_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0131_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0215_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0228_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0315_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0331_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0415_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0430_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0515_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0531_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0615_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0630_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0715_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0731_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0815_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0831_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0915_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "0930_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1015_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1031_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1115_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1130_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1215_" & iIndex & " />", "")
								sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & oRequest("YearID").Item & "1231_" & iIndex & " />", "")
							Next
						End If
						'sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_TOTAL_-2 />", FormatNumber(dDeductionsAmount, 2, True, False, True))
						'sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_TOTAL_-1 />", FormatNumber(dPerceptionsAmount, 2, True, False, True))
						'sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_TOTAL_0 />", FormatNumber((dPerceptionsAmount - dDeductionsAmount), 2, True, False, True))
						For iIndex = 0 To UBound(adTotalAmounts)
							sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & adTotalAmounts(iIndex)(0) & "_Count />", "<B>" & adTotalAmounts(iIndex)(1) & "</B>")
							sCurrentAmounts = Replace(sCurrentAmounts, "<CONCEPT_AMOUNT_" & adTotalAmounts(iIndex)(0) & "_Total />", "<B>" & FormatNumber(adTotalAmounts(iIndex)(2), 2, True, False, True) & "</B>")
						Next
						asRowContents = Split(sRowContents & sCurrentAmounts, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sFileName & ".xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
					lErrorNumber = AppendTextToFile(sFileName & ".xls", "</TABLE>", sErrorDescription)

					lErrorNumber = ZipFile(sFileName & ".xls", sFileName & ".zip", sErrorDescription)
					If lErrorNumber = 0 Then
						Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
						sErrorDescription = "No se pudo guardar la información del reporte."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
					If lErrorNumber = 0 Then
						lErrorNumber = DeleteFile(sFileName & ".xls", sErrorDescription)
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
	BuildReport1151 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1152(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paid amounts for every employee in
'         the given year for the DIM format
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1152"
	Dim bEmpty
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sTempFileName
	Dim sExcluded
	Dim sCondition
	Dim oRecordset
	Dim lCurrentID
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim jIndex
	Dim asTaxes
	Dim sFileContents
	Dim dTotal1
	Dim dTotal2
	Dim dTotal3
	Dim dTotal4
	Dim dTotal5
	Dim dTotal6
	Dim dTotal7
	Dim dTotal8
	Dim dTotal9
	Dim dTotal10
	Dim dTotal11
	Dim dTotal12
	Dim dTotal13
	Dim dTotal14
	Dim dSMG15
	Dim dSMG30
	Dim dISRFromTable
	Dim dRateFromTable
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bEmpty = True
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "(Companies.", "(Employees.")
	sCondition = Replace(sCondition, "(EmployeeTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(PositionTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(Journeys.", "(Employees.")
	sCondition = Replace(sCondition, "(Shifts.", "(Employees.")
	sCondition = Replace(sCondition, "(Levels.", "(Employees.")
	sCondition = Replace(sCondition, "(PaymentCenters.", "(Employees.")
	sCondition = Replace(sCondition, "(Positions.", "(Jobs.")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN)
	If Not FolderExists(Server.MapPath(sFilePath), sErrorDescription) Then lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFileName = sFilePath & "\Rep_"
	sFilePath = sFilePath & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFilePath = Server.MapPath(sFilePath)
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sTempFileName = Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "\Rep_")
		sExcluded = "-1"

		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, InvertedTax, InvertedRate From TaxInvertions Where (StartDate<=" & oRequest("YearID").Item & "0101) And (EndDate>=" & oRequest("YearID").Item & "1231) And (PeriodID=8)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asTaxes = ""
			Do While Not oRecordset.EOF
				asTaxes = asTaxes & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("InvertedTax").Value) & "," & CStr(oRecordset.Fields("InvertedRate").Value) & "" & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			asTaxes = asTaxes & "0,1E+69,0,1"
		End If
		asTaxes = Split(asTaxes, ";")
		For iIndex = 0 To UBound(asTaxes)
			asTaxes(iIndex) = Split(asTaxes(iIndex), ",")
			For jIndex = 0 To UBound(asTaxes(iIndex))
				asTaxes(iIndex)(jIndex) = CDbl(asTaxes(iIndex)(jIndex))
			Next
		Next

		lCurrentID = -1
		sErrorDescription = "No se pudieron obtener los registros de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, bTaxAdjustment, Areas.AreaShortName, Zones.ZoneCode, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue From Payroll_" & oRequest("YearID").Item & ", Concepts, Employees, EmployeesForTaxAdjustment, Jobs, Areas, Zones, ZoneTypes, CurrenciesHistoryList Where (Payroll_" & oRequest("YearID").Item & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesForTaxAdjustment.EmployeeID) And (EmployeesForTaxAdjustment.PayrollYear=" & oRequest("YearID").Item & ") And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (ZoneTypes.ZoneTypeID=CurrenciesHistoryList.CurrencyID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Concepts.ConceptID>0) And ((Concepts.IsDeduction=0) Or (Concepts.ConceptID In (30, 52, 55, 71, 72, 110))) And (CurrenciesHistoryList.CurrencyDate=" & oRequest("YearID").Item & "1231) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (Areas.EndDate=30000000) And (Zones.EndDate=30000000) " & sCondition & " Group By Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, bTaxAdjustment, Areas.AreaShortName, Zones.ZoneCode, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue Order By Employees.EmployeeID, Concepts.IsDeduction, Concepts.ConceptShortName", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sFileContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "Report_1152.htm"), sErrorDescription)
				If Len(sFileContents) > 0 Then
					sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					bEmpty = False
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							If lCurrentID > -1 Then
								dTotal2 = dTotal1 - dTotal2
								dTotal3 = dTotal2 - dTotal3
								dTotal4 = dTotal4 - dSMG30
								dTotal5 = dTotal3 + dTotal4
								For iIndex = 0 To UBound(asTaxes)
									If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
										dISRFromTable = asTaxes(iIndex)(2)
										dRateFromTable = asTaxes(iIndex)(3)
										Exit For
									End If
								Next
								dTotal6 = dTotal5 - dISRFromTable
								dTotal7 = dTotal6 / dRateFromTable
								dTotal8 = dTotal7 - dTotal2
								dTotal9 = dTotal8 + dSMG30
								dTotal10 = dTotal8 - dTotal4
								dTotal11 = dTotal1 + dTotal8
								dTotal12 = dTotal12 + dSMG15 + dSMG30
								dTotal13 = dTotal11 + dTotal12
								dTotal14 = dTotal14 + dTotal10
								sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
								For iIndex = 0 To 200
									sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & " />", "0.00")
									sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & "_SN />", "0.00")
									sRowContents = Replace(sRowContents, "<CONCEPT_TAX_" & iIndex & " />", "0.00")
								Next
								lErrorNumber = AppendTextToFile(sTempFileName & lCurrentID & ".doc", sRowContents, sErrorDescription)
							End If
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							sRowContents = ""
							dSMG15 = 0
							dSMG30 = 0
							dTotal1 = 0
							dTotal2 = 0
							dTotal3 = 0
							dTotal4 = 0
							dTotal5 = 0
							dTotal6 = 0
							dTotal7 = 0
							dTotal8 = 0
							dTotal9 = 0
							dTotal10 = 0
							dTotal11 = 0
							dTotal12 = 0
							dTotal13 = 0
							dTotal14 = 0
							sRowContents = sFileContents
							sRowContents = Replace(sRowContents, "<FIELD_1 />", "01")
							sRowContents = Replace(sRowContents, "<FIELD_2 />", "12")
							sRowContents = Replace(sRowContents, "<FIELD_3 />", CStr(oRecordset.Fields("RFC").Value))
							sRowContents = Replace(sRowContents, "<FIELD_4 />", CStr(oRecordset.Fields("CURP").Value))
							sRowContents = Replace(sRowContents, "<FIELD_5 />", Left(CStr(oRecordset.Fields("EmployeeLastName").Value), 43))
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = Replace(sRowContents, "<FIELD_6 />", Left(CStr(oRecordset.Fields("EmployeeLastName2").Value), 43))
							Else
								sRowContents = Replace(sRowContents, "<FIELD_6 />", "                                          ")
							End If
							sRowContents = Replace(sRowContents, "<FIELD_7 />", Left(CStr(oRecordset.Fields("EmployeeName").Value), 43))
							sRowContents = Replace(sRowContents, "<FIELD_8 />", Right(("00" & CStr(oRecordset.Fields("EconomicZoneID").Value)), Len("00")))
							sRowContents = Replace(sRowContents, "<FIELD_9 />", Replace(CStr(oRecordset.Fields("bTaxAdjustment").Value), "0", "2"))
							sRowContents = Replace(sRowContents, "<FIELD_10 />", "1")
							sRowContents = Replace(sRowContents, "<FIELD_11 />", "1")
							sRowContents = Replace(sRowContents, "<FIELD_12 />", "0.00")
							If InStr(1, sSyndicateIDs, "," & CStr(oRecordset.Fields("EmployeeID").Value) & ",", vbBinaryCompare) > 0 Then
								sRowContents = Replace(sRowContents, "<FIELD_13 />", "1")
							Else
								sRowContents = Replace(sRowContents, "<FIELD_13 />", "2")
							End If
							sRowContents = Replace(sRowContents, "<FIELD_14 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_15 />", CStr(oRecordset.Fields("ZoneCode").Value))
							sRowContents = Replace(sRowContents, "<FIELD_16 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_17 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_18 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_19 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_20 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_21 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_22 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_23 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_24 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_25 />", "")
							sRowContents = Replace(sRowContents, "<FIELD_26 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_27 />", "0")
							sRowContents = Replace(sRowContents, "<FIELD_28 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_29 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_30 />", "<CONCEPT_120_SN />")
							sRowContents = Replace(sRowContents, "<FIELD_31 />", "2")
							sRowContents = Replace(sRowContents, "<FIELD_32 />", "1")
							sRowContents = Replace(sRowContents, "<FIELD_33 />", "<TOTAL_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_34 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_35 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_36 />", "0.00")
							If (CInt(oRequest("YearID").Item) Mod 4) = 0 Then
								sRowContents = Replace(sRowContents, "<FIELD_37 />", "366")
							Else
								sRowContents = Replace(sRowContents, "<FIELD_37 />", "365")
							End If
							sRowContents = Replace(sRowContents, "<FIELD_38 />", "<TOTAL_3 />")
							sRowContents = Replace(sRowContents, "<FIELD_39 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_40 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_41 />", "<TOTAL_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_42 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_43 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_44 />", CalculateAgeFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), CLng(oRequest("YearID").Item & "1231")))
							sRowContents = Replace(sRowContents, "<FIELD_45 />", "<TOTAL_3 />")
							sRowContents = Replace(sRowContents, "<FIELD_46 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_47 />", "<TOTAL_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_48 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_49 />", "<TOTAL_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_50 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_51 />", "<CONCEPT_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_52 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_53 />", "2")
							sRowContents = Replace(sRowContents, "<FIELD_54 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_55 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_56 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_57 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_58 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_59 />", "<TOTAL_3 />")
							sRowContents = Replace(sRowContents, "<FIELD_60 />", "<TOTAL_4 />")
							sRowContents = Replace(sRowContents, "<FIELD_61 />", "<TOTAL_4 />")
							sRowContents = Replace(sRowContents, "<FIELD_62 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_63 />", "<CONCEPT_TAX_139 />")
							sRowContents = Replace(sRowContents, "<FIELD_64 />", "<CONCEPT_TAX_9 />")
							sRowContents = Replace(sRowContents, "<FIELD_65 />", "<CONCEPT_TAX_148 />")
							sRowContents = Replace(sRowContents, "<FIELD_66 />", "<CONCEPT_TAX_21 />")
							sRowContents = Replace(sRowContents, "<FIELD_67 />", "<CONCEPT_TAX_20 />")
							sRowContents = Replace(sRowContents, "<FIELD_68 />", "<CONCEPT_TAX_17 />")
							sRowContents = Replace(sRowContents, "<FIELD_69 />", "<CONCEPT_TAX_16 />")
							sRowContents = Replace(sRowContents, "<FIELD_70 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_71 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_72 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_73 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_74 />", "<CONCEPT_77 />")
							sRowContents = Replace(sRowContents, "<FIELD_75 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_76 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_77 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_78 />", "<CONCEPT_125 />")
							sRowContents = Replace(sRowContents, "<FIELD_79 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_80 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_81 />", "<CONCEPT_45 />")
							sRowContents = Replace(sRowContents, "<FIELD_82 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_83 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_84 />", "<CONCEPT_41 />")
							sRowContents = Replace(sRowContents, "<FIELD_85 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_86 />", "<CONCEPT_65 />")
							sRowContents = Replace(sRowContents, "<FIELD_87 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_88 />", "<CONCEPT_84 />")
							sRowContents = Replace(sRowContents, "<FIELD_89 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_90 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_91 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_92 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_93 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_94 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_95 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_96 />", "<CONCEPT_15 />")
							sRowContents = Replace(sRowContents, "<FIELD_97 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_98 />", "<CONCEPT_37 />")
							sRowContents = Replace(sRowContents, "<FIELD_99 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_100 />", "<CONCEPT_24 />")
							sRowContents = Replace(sRowContents, "<FIELD_101 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_102 />", "<CONCEPT_36 />")
							sRowContents = Replace(sRowContents, "<FIELD_103 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_104 />", "<CONCEPT_56 />")
							sRowContents = Replace(sRowContents, "<FIELD_105 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_106 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_107 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_108 />", "<CONCEPT_22 />")
							sRowContents = Replace(sRowContents, "<FIELD_109 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_110 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_111 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_112 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_113 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_114 />", "<TOTAL_2 />")
							sRowContents = Replace(sRowContents, "<FIELD_115 />", "<TOTAL_3 />")
							sRowContents = Replace(sRowContents, "<FIELD_116 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_117 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_118 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_119 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_120 />", "<CONCEPT_48 />")
							sRowContents = Replace(sRowContents, "<FIELD_121 />", "<CONCEPT_49 />")
							sRowContents = Replace(sRowContents, "<FIELD_122 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_123 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_124 />", "<CONCEPT_1 />")
							sRowContents = Replace(sRowContents, "<FIELD_125 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_126 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_127 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_128 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_129 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_130 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_131 />", "<CONCEPT_55 />")
							sRowContents = Replace(sRowContents, "<FIELD_132 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_133 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_134 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_135 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_136 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_137 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_138 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_139 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_140 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_141 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_142 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_143 />", "0.00")
							sRowContents = Replace(sRowContents, "<FIELD_144 />", CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = Replace(sRowContents, "<FIELD_145 />", CStr(oRecordset.Fields("AreaShortName").Value))
							sRowContents = Replace(sRowContents, "<FIELD_146 />", "1")
						End If
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 20 '18 Prima de vacaciones exenta
'								sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />", 1, -1, vbBinaryCompare)
'								dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dSMG15 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 15
								If dSMG15 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG15 = CDbl(oRecordset.Fields("TotalAmount").Value)
'							Case 21 '18 Prima de vacaciones gravable
'								sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />")
'								dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 30 '26. Aguinaldo
								sRowContents = Replace(sRowContents, "<CONCEPT_30 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dSMG30 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 30
								If dSMG30 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG30 = CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 44, 94 '41 Premio antigüedad 25 y 30 años (mes de sueldo), C3 Premios, estimulos y recompensas (recompensa del sistema de evaluación del desempeño)
								sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><PIRAMID_CONCEPTS />")
								dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 52, 71, 72 '50 Faltas, 70 Retardos, 71 Deducción por cobro de sueldos indebidos
								sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><DEDUCTION_CONCEPTS />")
								dTotal2 = dTotal2 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 55 '53 Impuesto sobre producto de trabajo (ISR)
								sRowContents = Replace(sRowContents, "<CONCEPT_55 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 110 'IS ISR patronal del Seguro de Separación Individualizado
								sRowContents = Replace(sRowContents, "<CONCEPT_110 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case Else
								If (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) > 0) Then
									sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
									dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
								ElseIf (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) = 0) Then
									sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
									dTotal12 = dTotal12 + CDbl(oRecordset.Fields("TotalAmount").Value)
								End If
						End Select
						sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & "_SN />", "1")
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					dTotal2 = dTotal1 - dTotal2
					dTotal3 = dTotal2 - dTotal3
					dTotal4 = dTotal4 - dSMG30
					dTotal5 = dTotal3 + dTotal4
					For iIndex = 0 To UBound(asTaxes)
						If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
							dISRFromTable = asTaxes(iIndex)(2)
							dRateFromTable = asTaxes(iIndex)(3)
							Exit For
						End If
					Next
					dTotal6 = dTotal5 - dISRFromTable
					dTotal7 = dTotal6 / dRateFromTable
					dTotal8 = dTotal7 - dTotal2
					dTotal9 = dTotal8 + dSMG30
					dTotal10 = dTotal8 - dTotal4
					dTotal11 = dTotal1 + dTotal8
					dTotal12 = dTotal12 + dSMG15 + dSMG30
					dTotal13 = dTotal11 + dTotal12
					dTotal14 = dTotal14 + dTotal10
					sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
					For iIndex = 0 To 200
						sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & " />", "0.00")
						sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & "_SN />", "0.00")
						sRowContents = Replace(sRowContents, "<CONCEPT_TAX_" & iIndex & " />", "0.00")
					Next
					lErrorNumber = AppendTextToFile(sTempFileName & lCurrentID & ".doc", sRowContents, sErrorDescription)
				End If
				oRecordset.Close
			End If
		End If

		If Not bEmpty Then
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudo guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1152 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1153(oRequest, oADODBConnection, bForRecord, sErrorDescription)
'************************************************************
'Purpose: To display the paid amounts for every employee in
'         the given year
'Inputs:  oRequest, oADODBConnection, bForRecord
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1153"
	Dim sDate
	Dim sFileName
	Dim sExcluded
	Dim sCondition
	Dim oRecordset
	Dim dMaxDiscount
	Dim dLimit1
	Dim dLimit2
	Dim lCurrentID
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim jIndex
	Dim dAmount
	Dim dTaxAmount
	Dim dTemp
	Dim asDSM
	Dim asTaxes
	Dim asAllowances
	Dim sFileContents
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "(Companies.", "(Employees.")
	sCondition = Replace(sCondition, "(EmployeeTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(PositionTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(Journeys.", "(Employees.")
	sCondition = Replace(sCondition, "(Shifts.", "(Employees.")
	sCondition = Replace(sCondition, "(Levels.", "(Employees.")
	sCondition = Replace(sCondition, "(PaymentCenters.", "(Employees.")
	sCondition = Replace(sCondition, "(Positions.", "(Jobs.")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN)
	If Not FolderExists(Server.MapPath(sFileName), sErrorDescription) Then lErrorNumber = CreateFolder(Server.MapPath(sFileName), sErrorDescription)
	sFileName = sFileName & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & ".zip" & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sFileName = Server.MapPath(sFileName)
		sExcluded = "-1"

		sErrorDescription = "No se pudieron obtener los registros de los empleados que no estuvieron activos los 365 días del año."
		'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName As ZoneName02, Zones01.ZoneName As ZoneName01, GeneratingAreas.GeneratingAreaName From Employees, EmployeesHistoryList, StatusEmployees, Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And ((Employees.Active=0) Or (EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0)) And (EmployeesHistoryList.EmployeeDate>" & oRequest("YearID").Item & "0000) And (EmployeesHistoryList.EmployeeDate<" & oRequest("YearID").Item & "9999) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP From Employees, EmployeesHistoryList, StatusEmployees, Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And ((Employees.Active=0) Or (EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0)) And (EmployeesHistoryList.EmployeeDate>" & oRequest("YearID").Item & "0000) And (EmployeesHistoryList.EmployeeDate<" & oRequest("YearID").Item & "9999) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					sExcluded = sExcluded & "," & CStr(oRecordset.Fields("EmployeeID").Value)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
			oRecordset.Close
		End If

		dMaxDiscount = 9E+100
		dLimit1 = -9E+100
		dLimit2 = 9E+100
		If bForRecord Then
			If Len(oRequest("MaxDiscount").Item) > 0 Then dMaxDiscount = CDbl(oRequest("MaxDiscount").Item)
			If Len(oRequest("Limit1").Item) > 0 Then dLimit1 = CDbl(oRequest("Limit1").Item)
			If Len(oRequest("Limit2").Item) > 0 Then dLimit2 = CDbl(oRequest("Limit2").Item)
		End If
		sErrorDescription = "No se pudieron obtener los registros de los empleados cuyos ingresos son mayores a 400,000."
'		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName As ZoneName02, Zones01.ZoneName As ZoneName01, GeneratingAreas.GeneratingAreaName, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Employees, Payroll_" & oRequest("YearID").Item & ", Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=0) " & sCondition & " Group by Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode, PaymentCenters.AreaName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName, Zones01.ZoneName, GeneratingAreas.GeneratingAreaName Order By Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Employees, Payroll_" & oRequest("YearID").Item & ", Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=0) " & sCondition & " Group by Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP Order By Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					If CDbl(oRecordset.Fields("TotalAmount").Value) <= dLimit2 Then Exit Do
					sExcluded = sExcluded & "," & CStr(oRecordset.Fields("EmployeeID").Value)
					If bForRecord Then
						lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, 0", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
			oRecordset.Close
		End If

		sErrorDescription = "No se pudo obtener la lista de empleados que no quieren desean el ajuste anual del impuesto sobre la renta."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP From Employees, EmployeesForTaxAdjustment, Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=EmployeesForTaxAdjustment.EmployeeID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (EmployeesForTaxAdjustment.PayrollYear=" & oRequest("YearID").Item & ") And (EmployeesForTaxAdjustment.bTaxAdjustment=0) " & sCondition & " Order By Employees.EmployeeNumber", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					sExcluded = sExcluded & "," & CStr(oRecordset.Fields("EmployeeID").Value)
					If bForRecord Then
						lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, 0", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
			oRecordset.Close
		End If

		sErrorDescription = "No se pudieron obtener los días de salario mínimo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyDate=" & oRequest("YearID").Item & "1231) And (CurrencyID In (1,2,3,4,5)) Order By CurrencyID", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asDSM = "0;"
			Do While Not oRecordset.EOF
				asDSM = asDSM & CStr(oRecordset.Fields("CurrencyValue").Value) & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		End If
		asDSM = Split(asDSM, ";")
		For iIndex = 0 To UBound(asDSM)
			asDSM(iIndex) = CDbl(asDSM(iIndex))
		Next

		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, FixedAmount, PercentageForExcess From TaxLimits Where (StartDate<=" & oRequest("YearID").Item & "1231) And (EndDate>=" & oRequest("YearID").Item & "1231) And (PeriodID=4)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asTaxes = ""
			Do While Not oRecordset.EOF
				asTaxes = asTaxes & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("FixedAmount").Value) & "," & CStr(oRecordset.Fields("PercentageForExcess").Value) & "" & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			asTaxes = asTaxes & "0,1E+69,0,1"
		End If
		asTaxes = Split(asTaxes, ";")
		For iIndex = 0 To UBound(asTaxes)
			asTaxes(iIndex) = Split(asTaxes(iIndex), ",")
			For jIndex = 0 To UBound(asTaxes(iIndex))
				asTaxes(iIndex)(jIndex) = CDbl(asTaxes(iIndex)(jIndex))
			Next
		Next

		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, AllowanceAmount From EmploymentAllowances Where (StartDate<=" & oRequest("YearID").Item & "1231) And (EndDate>=" & oRequest("YearID").Item & "1231) And (PeriodID=4)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asAllowances = ""
			Do While Not oRecordset.EOF
				asAllowances = asAllowances & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("AllowanceAmount").Value) & "" & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			asAllowances = asAllowances & "0,1E+69,0"
		End If
		asAllowances = Split(asAllowances, ";")
		For iIndex = 0 To UBound(asAllowances)
			asAllowances(iIndex) = Split(asAllowances(iIndex), ",")
			For jIndex = 0 To UBound(asAllowances(iIndex))
				asAllowances(iIndex)(jIndex) = CDbl(asAllowances(iIndex)(jIndex))
			Next
		Next

		sErrorDescription = "No se pudieron obtener los registros de los empleados que estuvieron activos los 365 días del año."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Companies.CompanyShortName, Companies.CompanyName, Areas.AreaShortName, Areas.EconomicZoneID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Sum(ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", Concepts, Employees, Companies, Jobs, Areas, Zones Where (Payroll_" & oRequest("YearID").Item & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (Employees.CompanyID=Companies.CompanyID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (RecordDate<" & oRequest("YearID").Item & "9999) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount=100)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount=0)) Or (Concepts.ConceptID=55)) And (Concepts.ConceptID>0) " & sCondition & " Group By Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Companies.CompanyShortName, Companies.CompanyName, Areas.AreaShortName, Areas.EconomicZoneID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount Order By Employees.EmployeeNumber", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			lErrorNumber = AppendTextToFile(sFileName & ".xls", "<TABLE BORDER=""1"">", sErrorDescription)
				'sRowContents = "<TR><TD ALIGN=""CENTER""><B>No. del empleado</B></TD><TD ALIGN=""CENTER""><B>Apellido paterno</B></TD><TD ALIGN=""CENTER""><B>Apellido materno</B></TD><TD ALIGN=""CENTER""><B>Nombre</B></TD><TD ALIGN=""CENTER""><B>RFC</B></TD><TD ALIGN=""CENTER""><B>CURP</B></TD><TD ALIGN=""CENTER""><B>Base gravable</B></TD><TD ALIGN=""CENTER""><B>Límite inferior</B></TD><TD ALIGN=""CENTER""><B>Restante</B></TD><TD ALIGN=""CENTER""><B>% excendente</B></TD><TD ALIGN=""CENTER""><B>Cuota fija</B></TD><TD ALIGN=""CENTER""><B>Subsidio</B></TD><TD ALIGN=""CENTER""><B>Impuesto</B></TD><TD ALIGN=""CENTER""><B>Impuesto retenido en todo el año</B></TD><TD ALIGN=""CENTER""><B>Ajuste</B></TD></TR>"
				sRowContents = "<TR><TD ALIGN=""CENTER""><B>No. del empleado</B></TD><TD ALIGN=""CENTER""><B>Apellido paterno</B></TD><TD ALIGN=""CENTER""><B>Apellido materno</B></TD><TD ALIGN=""CENTER""><B>Nombre</B></TD><TD ALIGN=""CENTER""><B>RFC</B></TD><TD ALIGN=""CENTER""><B>CURP</B></TD><TD ALIGN=""CENTER""><B>Empresa</B></TD><TD ALIGN=""CENTER""><B>Centro de trabajo</B></TD><TD ALIGN=""CENTER""><B>Salario mínimo general</B></TD><TD ALIGN=""CENTER""><B>Concepto</B></TD><TD ALIGN=""CENTER""><B>Monto</B></TD>"
				lErrorNumber = AppendTextToFile(sFileName & ".xls", sRowContents, sErrorDescription)

				lCurrentID = -1
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID <> -1 Then
							'sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dAmount, 2, True, False, True)
							sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Base gravable</B><BR /><CONCEPT_NAME />")
							sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
							dTemp = dAmount
							For iIndex = 0 To UBound(asTaxes)
								If (asTaxes(iIndex)(0) <= dAmount) And (asTaxes(iIndex)(1) >= dAmount) Then
									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Límite inferior<BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(0), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Restante<BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber((dAmount - asTaxes(iIndex)(0)), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;% excendente<BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(3), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Cuota fija<BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(2), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

									dAmount = ((dAmount - asTaxes(iIndex)(0)) * asTaxes(iIndex)(3) / 100) + asTaxes(iIndex)(2)
									Exit For
								End If
							Next
							For iIndex = 0 To UBound(asAllowances)
								If (asAllowances(iIndex)(0) <= dTemp) And (asAllowances(iIndex)(1) >= dTemp) Then
									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Subsidio<BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asAllowances(iIndex)(2), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

									dAmount = dAmount - asAllowances(iIndex)(2)
									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Impuesto</B><BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")

									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Impuesto retenido en todo el año</B><BR /><CONCEPT_NAME />")
									sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dTaxAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")

									sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Ajuste</B><BR /><CONCEPT_NAME />")
									If bForRecord Then
'										If (dLimit1 >= CDbl(FormatNumber((dAmount - dTaxAmount), 2, True, False, True))) Or (dLimit2 <= CDbl(FormatNumber((dAmount - dTaxAmount), 2, True, False, True))) Then
'											sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</FONT><BR /><CONCEPT_VALUE />")
'											lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, 0", sErrorDescription)
'										Else
'											sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
'											lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, " & dAmount - dTaxAmount, sErrorDescription)
'										End If
										If dMaxDiscount < CDbl(FormatNumber((dAmount - dTaxAmount), 2, True, False, True)) Then
											sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dMaxDiscount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
											lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, " & dMaxDiscount, sErrorDescription)
										Else
											sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
											lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, " & dAmount - dTaxAmount, sErrorDescription)
										End If
									Else
										sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
									End If
									Exit For
								End If
							Next
							asRowContents = Split(sRowContents, TABLE_SEPARATOR)
							lErrorNumber = AppendTextToFile(sFileName & ".xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							dAmount = 0
							dTaxAmount = 0
						End If
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & " "
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asDSM(CInt(oRecordset.Fields("EconomicZoneID").Value)), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_NAME />" & TABLE_SEPARATOR & "<CONCEPT_VALUE />"
					End If
					If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". "& CStr(oRecordset.Fields("ConceptName").Value)) & "<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")
						dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
					ElseIf CDbl(oRecordset.Fields("TaxAmount").Value) = 0 Then
						dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". "& CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT><BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</FONT><BR /><CONCEPT_VALUE />")
					Else
						dTaxAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				'sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(dAmount, 2, True, False, True)
				sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Base gravable</B><BR /><CONCEPT_NAME />")
				sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
				dTemp = dAmount
				For iIndex = 0 To UBound(asTaxes)
					If (asTaxes(iIndex)(0) <= dAmount) And (asTaxes(iIndex)(1) >= dAmount) Then
						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Límite inferior<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(0), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Restante<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber((dAmount - asTaxes(iIndex)(0)), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;% excendente<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(3), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Cuota fija<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asTaxes(iIndex)(2), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

						dAmount = ((dAmount - asTaxes(iIndex)(0)) * asTaxes(iIndex)(3) / 100) + asTaxes(iIndex)(2)
						Exit For
					End If
				Next
				For iIndex = 0 To UBound(asAllowances)
					If (asAllowances(iIndex)(0) <= dTemp) And (asAllowances(iIndex)(1) >= dTemp) Then
						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "&nbsp;&nbsp;&nbsp;Subsidio<BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", FormatNumber(asAllowances(iIndex)(2), 2, True, False, True) & "<BR /><CONCEPT_VALUE />")

						dAmount = dAmount - asAllowances(iIndex)(2)
						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Impuesto</B><BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")

						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Impuesto retenido en todo el año</B><BR /><CONCEPT_NAME />")
						sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber(dTaxAmount, 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")

						sRowContents = Replace(sRowContents, "<CONCEPT_NAME />", "<B>Ajuste</B><BR /><CONCEPT_NAME />")
						If bForRecord Then
							If (dLimit1 >= (dAmount - dTaxAmount)) Or (dLimit2 <= (dAmount - dTaxAmount)) Then
								sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</FONT><BR /><CONCEPT_VALUE />")
								lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, 0", sErrorDescription)
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_VALUE />", "<B>" & FormatNumber((dAmount - dTaxAmount), 2, True, False, True) & "</B><BR /><CONCEPT_VALUE />")
								lErrorNumber = AppendTextToFile(sFileName & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", 55, 1, " & dAmount - dTaxAmount, sErrorDescription)
							End If
						End If
						Exit For
					End If
				Next
				asRowContents = Split(sRowContents, TABLE_SEPARATOR)
				lErrorNumber = AppendTextToFile(sFileName & ".xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
			lErrorNumber = AppendTextToFile(sFileName & ".xls", "</TABLE>", sErrorDescription)
		End If
		If bForRecord Then
			sRowContents = GetFileContents(sFileName & ".txt", sErrorDescription)
			If Len(sRowContents) > 0 Then
				sErrorDescription = "No se pudo actualizar la información del ajuste anual."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & oRequest("YearID").Item & " Where (RecordDate=" & oRequest("YearID").Item & "9999)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo actualizar la información del ajuste anual."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & oRequest("YearID").Item & "1231 Where (RecordID=99) And (ConceptID=55)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					asRowContents = Split(sRowContents, vbNewLine)
					For iIndex = 0 To UBound(asRowContents)
						If Len(asRowContents(iIndex)) > 0 Then
							sErrorDescription = "No se pudo guardar la información del ajuste anual."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & oRequest("YearID").Item & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & oRequest("YearID").Item & "9999, 1, " & asRowContents(iIndex) & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo guardar la información del ajuste anual."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & oRequest("YearID").Item & "1231 (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & oRequest("YearID").Item & "1231, 99, " & asRowContents(iIndex) & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
						End If
					Next
					lErrorNumber = DeleteFile(sFileName & ".txt", sErrorDescription)
				End If
			End If
		End If

		lErrorNumber = ZipFile(sFileName & ".xls", sFileName & ".zip", sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudo guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		If lErrorNumber = 0 Then
			lErrorNumber = DeleteFile(sFileName & ".xls", sErrorDescription)
		End If
		oEndDate = Now()
		If (lErrorNumber = 0) And B_USE_SMTP Then
			If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1153 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1154(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paid amounts for every employee in
'         the given year
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1154"
	Dim bEmpty
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sTempFileName
	Dim sExcluded
	Dim sCondition
	Dim oRecordset
	Dim lCurrentID
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim jIndex
	Dim asTaxes
	Dim sFileContents
	Dim dTotal1
	Dim dTotal2
	Dim dTotal3
	Dim dTotal4
	Dim dTotal5
	Dim dTotal6
	Dim dTotal7
	Dim dTotal8
	Dim dTotal9
	Dim dTotal10
	Dim dTotal11
	Dim dTotal12
	Dim dTotal13
	Dim dTotal14
	Dim dSMG15
	Dim dSMG30
	Dim dISRFromTable
	Dim dRateFromTable
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bEmpty = True
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "(Companies.", "(Employees.")
	sCondition = Replace(sCondition, "(EmployeeTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(PositionTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(Journeys.", "(Employees.")
	sCondition = Replace(sCondition, "(Shifts.", "(Employees.")
	sCondition = Replace(sCondition, "(Levels.", "(Employees.")
	sCondition = Replace(sCondition, "(PaymentCenters.", "(Employees.")
	sCondition = Replace(sCondition, "(Positions.", "(Jobs.")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN)
	If Not FolderExists(Server.MapPath(sFilePath), sErrorDescription) Then lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFileName = sFilePath & "\Rep_"
	sFilePath = sFilePath & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFilePath = Server.MapPath(sFilePath)
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sTempFileName = Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "\Rep_")
		sExcluded = "-1"

		sErrorDescription = "No se pudieron obtener los registros de los empleados que no estuvieron activos los 365 días del año."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName As ZoneName02, Zones01.ZoneName As ZoneName01, GeneratingAreas.GeneratingAreaName From Employees, EmployeesHistoryList, StatusEmployees, Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And ((Employees.Active=0) Or (EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0)) And (EmployeesHistoryList.EmployeeDate>" & oRequest("YearID").Item & "0000) And (EmployeesHistoryList.EmployeeDate<" & oRequest("YearID").Item & "9999) " & sCondition & " Order By Employees.EmployeeID", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				lErrorNumber = AppendTextToFile(sTempFileName & "Inactive.xls", "<B>EMPLEADOS QUE NO ESTUVIERON ACTIVOS LOS 365 DÍAS DEL AÑO</B><BR /><BR />", sErrorDescription)
				lErrorNumber = AppendTextToFile(sTempFileName & "Inactive.xls", "<TABLE BORDER=""1"">", sErrorDescription)
					sRowContents = "<TR><TD ALIGN=""CENTER""><B>No. del empleado</B></TD><TD ALIGN=""CENTER""><B>Apellido paterno</B></TD><TD ALIGN=""CENTER""><B>Apellido materno</B></TD><TD ALIGN=""CENTER""><B>Nombre</B></TD><TD ALIGN=""CENTER""><B>RFC</B></TD><TD ALIGN=""CENTER""><B>CURP</B></TD><TD ALIGN=""CENTER""><B>Fecha de ingreso</B></TD><TD ALIGN=""CENTER""><B>Tipo de puesto</B></TD><TD ALIGN=""CENTER""><B>Tipo de tabulador</B></TD><TD ALIGN=""CENTER""><B>Jornada</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Turno</B></TD><TD ALIGN=""CENTER""><B>Hora de entrada</B></TD><TD ALIGN=""CENTER""><B>Hora de salida</B></TD><TD ALIGN=""CENTER""><B>Hora de entrada</B></TD><TD ALIGN=""CENTER""><B>Hora de salida</B></TD><TD ALIGN=""CENTER""><B>Horas laboradas</B></TD><TD ALIGN=""CENTER""><B>Número de seguro social</B></TD><TD ALIGN=""CENTER""><B>Plaza</B></TD><TD ALIGN=""CENTER""><B>Zona económica</B></TD><TD ALIGN=""CENTER""><B>Puesto</B></TD><TD ALIGN=""CENTER""><B>Nivel-subnivel</B></TD><TD ALIGN=""CENTER""><B>Grupo, grado, nivel</B></TD><TD ALIGN=""CENTER""><B>Integración</B></TD><TD ALIGN=""CENTER""><B>Clasificación</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Centro de pago</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Servicio</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Centro de trabajo</B></TD><TD ALIGN=""CENTER""><B>Entidad</B></TD><TD ALIGN=""CENTER""><B>Municipio</B></TD><TD ALIGN=""CENTER""><B>Población</B></TD><TD ALIGN=""CENTER""><B>Área generadora</B></TD></TR>"
					lErrorNumber = AppendTextToFile(sTempFileName & "Inactive.xls", sRowContents, sErrorDescription)
					Do While Not oRecordset.EOF
						sExcluded = sExcluded & "," & CStr(oRecordset.Fields("EmployeeID").Value)
						sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & " "
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplaydateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("StartHour1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("EndHour1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("StartHour2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("EndHour2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneTypeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName01").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName02").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GeneratingAreaName").Value))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sTempFileName & "Inactive.xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				lErrorNumber = AppendTextToFile(sTempFileName & "Inactive.xls", "</TABLE>", sErrorDescription)
			End If
			oRecordset.Close
		End If

		sErrorDescription = "No se pudieron obtener los registros de los empleados cuyos ingresos son mayores a 400,000."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName As ZoneName02, Zones01.ZoneName As ZoneName01, GeneratingAreas.GeneratingAreaName, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Employees, Payroll_" & oRequest("YearID").Item & ", Services, EmployeeTypes, PositionTypes, GroupGradeLevels, Journeys, Shifts, Levels, Areas As PaymentCenters, Jobs, Zones, Zones As Zones01, Zones As Zones02, ZoneTypes, Areas, GeneratingAreas, Positions Where (Employees.EmployeeID=Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (Employees.ServiceID=Services.ServiceID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (Zones.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=0) " & sCondition & " Group by Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, Jobs.JobNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelName, GroupGradeLevels.GroupGradeLevelName, Employees.IntegrationID, Employees.ClassificationID, PaymentCenters.AreaCode, PaymentCenters.AreaName, Services.ServiceShortName, Services.ServiceName, Areas.AreaCode, Areas.AreaName, Zones.ZoneName, Zones02.ZoneName, Zones01.ZoneName, GeneratingAreas.GeneratingAreaName Order By Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) Desc", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				lErrorNumber = AppendTextToFile(sTempFileName & "Amount.xls", "<B>EMPLEADOS QUE RECIBIERON INGRESOS MAYORES A $400,000.00</B><BR /><BR />", sErrorDescription)
				lErrorNumber = AppendTextToFile(sTempFileName & "Amount.xls", "<TABLE BORDER=""1"">", sErrorDescription)
					sRowContents = "<TR><TD ALIGN=""CENTER""><B>No. del empleado</B></TD><TD ALIGN=""CENTER""><B>Apellido paterno</B></TD><TD ALIGN=""CENTER""><B>Apellido materno</B></TD><TD ALIGN=""CENTER""><B>Nombre</B></TD><TD ALIGN=""CENTER""><B>RFC</B></TD><TD ALIGN=""CENTER""><B>CURP</B></TD><TD ALIGN=""CENTER""><B>Fecha de ingreso</B></TD><TD ALIGN=""CENTER""><B>Tipo de puesto</B></TD><TD ALIGN=""CENTER""><B>Tipo de tabulador</B></TD><TD ALIGN=""CENTER""><B>Jornada</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Turno</B></TD><TD ALIGN=""CENTER""><B>Hora de entrada</B></TD><TD ALIGN=""CENTER""><B>Hora de salida</B></TD><TD ALIGN=""CENTER""><B>Hora de entrada</B></TD><TD ALIGN=""CENTER""><B>Hora de salida</B></TD><TD ALIGN=""CENTER""><B>Horas laboradas</B></TD><TD ALIGN=""CENTER""><B>Número de seguro social</B></TD><TD ALIGN=""CENTER""><B>Plaza</B></TD><TD ALIGN=""CENTER""><B>Zona económica</B></TD><TD ALIGN=""CENTER""><B>Puesto</B></TD><TD ALIGN=""CENTER""><B>Nivel-subnivel</B></TD><TD ALIGN=""CENTER""><B>Grupo, grado, nivel</B></TD><TD ALIGN=""CENTER""><B>Integración</B></TD><TD ALIGN=""CENTER""><B>Clasificación</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Centro de pago</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Servicio</B></TD><TD ALIGN=""CENTER"" COLSPAN=""2""><B>Centro de trabajo</B></TD><TD ALIGN=""CENTER""><B>Entidad</B></TD><TD ALIGN=""CENTER""><B>Municipio</B></TD><TD ALIGN=""CENTER""><B>Población</B></TD><TD ALIGN=""CENTER""><B>Área generadora</B></TD></TR>"
					lErrorNumber = AppendTextToFile(sTempFileName & "Amount.xls", sRowContents, sErrorDescription)
					Do While Not oRecordset.EOF
						If CDbl(oRecordset.Fields("TotalAmount").Value) <= 400000 Then Exit Do
						sExcluded = sExcluded & "," & CStr(oRecordset.Fields("EmployeeID").Value)
						sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & " "
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplaydateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("StartHour1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("EndHour1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("StartHour2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayTimeFromSerialNumber(CStr(oRecordset.Fields("EndHour2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneTypeName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName01").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName02").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GeneratingAreaName").Value))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = AppendTextToFile(sTempFileName & "Amount.xls", GetTableRowText(asRowContents, True, sErrorDescription), sErrorDescription)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				lErrorNumber = AppendTextToFile(sTempFileName & "Amount.xls", "</TABLE>", sErrorDescription)
			End If
			oRecordset.Close
		End If

		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, InvertedTax, InvertedRate From TaxInvertions Where (StartDate<=" & oRequest("YearID").Item & "0101) And (EndDate>=" & oRequest("YearID").Item & "1231) And (PeriodID=8)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asTaxes = ""
			Do While Not oRecordset.EOF
				asTaxes = asTaxes & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("InvertedTax").Value) & "," & CStr(oRecordset.Fields("InvertedRate").Value) & "" & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			asTaxes = asTaxes & "0,1E+69,0,1"
		End If
		asTaxes = Split(asTaxes, ";")
		For iIndex = 0 To UBound(asTaxes)
			asTaxes(iIndex) = Split(asTaxes(iIndex), ",")
			For jIndex = 0 To UBound(asTaxes(iIndex))
				asTaxes(iIndex)(jIndex) = CDbl(asTaxes(iIndex)(jIndex))
			Next
		Next

		lCurrentID = -1
		sErrorDescription = "No se pudieron obtener los registros de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue From Payroll_" & oRequest("YearID").Item & ", Concepts, Employees, Jobs, Areas, Zones, ZoneTypes, CurrenciesHistoryList Where (Payroll_" & oRequest("YearID").Item & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (ZoneTypes.ZoneTypeID=CurrenciesHistoryList.CurrencyID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Concepts.ConceptID>0) And ((Concepts.IsDeduction=0) Or (Concepts.ConceptID In (30, 52, 55, 71, 72, 110))) And (CurrenciesHistoryList.CurrencyDate=" & oRequest("YearID").Item & "1231) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) " & sCondition & " Group By Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue Order By Employees.EmployeeID, Concepts.IsDeduction, Concepts.ConceptShortName", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sFileContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "Report_1154.htm"), sErrorDescription)
				If Len(sFileContents) > 0 Then
					bEmpty = False
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							If lCurrentID > -1 Then
								dTotal2 = dTotal1 - dTotal2
								dTotal3 = dTotal2 - dTotal3
								dTotal4 = dTotal4 - dSMG30
								dTotal5 = dTotal3 + dTotal4
								For iIndex = 0 To UBound(asTaxes)
									If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
										dISRFromTable = asTaxes(iIndex)(2)
										dRateFromTable = asTaxes(iIndex)(3)
										Exit For
									End If
								Next
								dTotal6 = dTotal5 - dISRFromTable
								dTotal7 = dTotal6 / dRateFromTable
								dTotal8 = dTotal7 - dTotal2
								dTotal9 = dTotal8 + dSMG30
								dTotal10 = dTotal8 - dTotal4
								dTotal11 = dTotal1 + dTotal8
								dTotal12 = dTotal12 + dSMG15 + dSMG30
								dTotal13 = dTotal11 + dTotal12
								dTotal14 = dTotal14 + dTotal10
								sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
								sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
								sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
								lErrorNumber = AppendTextToFile(sTempFileName & lCurrentID & ".xls", sRowContents, sErrorDescription)
							End If
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							sRowContents = sFileContents
							dSMG15 = 0
							dSMG30 = 0
							dTotal1 = 0
							dTotal2 = 0
							dTotal3 = 0
							dTotal4 = 0
							dTotal5 = 0
							dTotal6 = 0
							dTotal7 = 0
							dTotal8 = 0
							dTotal9 = 0
							dTotal10 = 0
							dTotal11 = 0
							dTotal12 = 0
							dTotal13 = 0
							dTotal14 = 0
							sRowContents = Replace(sRowContents, "<YEAR />", oRequest("YearID").Item)
							sRowContents = Replace(sRowContents, "<LAST_YEAR />", (CInt(oRequest("YearID").Item) - 1))
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)))
							Else
								sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)))
							End If
							sRowContents = Replace(sRowContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
							sRowContents = Replace(sRowContents, "<ZONE_TYPE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ZoneTypeName").Value)))
						End If
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 20 '18 Prima de vacaciones exenta
'								sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />", 1, -1, vbBinaryCompare)
'								dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dSMG15 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 15
								If dSMG15 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG15 = CDbl(oRecordset.Fields("TotalAmount").Value)
'							Case 21 '18 Prima de vacaciones gravable
'								sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />")
'								dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 30 '26. Aguinaldo
								sRowContents = Replace(sRowContents, "<CONCEPT_30 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dSMG30 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 30
								If dSMG30 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG30 = CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 44, 94 '41 Premio antigüedad 25 y 30 años (mes de sueldo), C3 Premios, estimulos y recompensas (recompensa del sistema de evaluación del desempeño)
								sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><PIRAMID_CONCEPTS />")
								dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 52, 71, 72 '50 Faltas, 70 Retardos, 71 Deducción por cobro de sueldos indebidos
								sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><DEDUCTION_CONCEPTS />")
								dTotal2 = dTotal2 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 55 '53 Impuesto sobre producto de trabajo (ISR)
								sRowContents = Replace(sRowContents, "<CONCEPT_55 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case 110 'IS ISR patronal del Seguro de Separación Individualizado
								sRowContents = Replace(sRowContents, "<CONCEPT_110 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
								dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
							Case Else
								If (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) > 0) Then
									sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />")
									dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
								ElseIf (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) = 0) Then
									sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><EXEMPT_CONCEPTS />")
									dTotal12 = dTotal12 + CDbl(oRecordset.Fields("TotalAmount").Value)
								End If
						End Select
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					dTotal2 = dTotal1 - dTotal2
					dTotal3 = dTotal2 - dTotal3
					dTotal4 = dTotal4 - dSMG30
					dTotal5 = dTotal3 + dTotal4
					For iIndex = 0 To UBound(asTaxes)
						If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
							dISRFromTable = asTaxes(iIndex)(2)
							dRateFromTable = asTaxes(iIndex)(3)
							Exit For
						End If
					Next
					dTotal6 = dTotal5 - dISRFromTable
					dTotal7 = dTotal6 / dRateFromTable
					dTotal8 = dTotal7 - dTotal2
					dTotal9 = dTotal8 + dSMG30
					dTotal10 = dTotal8 - dTotal4
					dTotal11 = dTotal1 + dTotal8
					dTotal12 = dTotal12 + dSMG15 + dSMG30
					dTotal13 = dTotal11 + dTotal12
					dTotal14 = dTotal14 + dTotal10
					sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
					sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
					sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
					lErrorNumber = AppendTextToFile(sTempFileName & lCurrentID & ".xls", sRowContents, sErrorDescription)
				End If
				oRecordset.Close
			End If
		End If

		If Not bEmpty Then
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudo guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1154 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1157(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the paid amounts for every employee in
'         the given year for the DIM format
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1157"
	Dim bEmpty
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sTempFileName
	Dim sExcluded
	Dim sCondition
	Dim oRecordset
	Dim lCurrentID
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim iIndex
	Dim jIndex
	Dim asTaxes
	Dim sFileContents
	Dim dTotal1
	Dim dTotal2
	Dim dTotal3
	Dim dTotal4
	Dim dTotal5
	Dim dTotal6
	Dim dTotal7
	Dim dTotal8
	Dim dTotal9
	Dim dTotal10
	Dim dTotal11
	Dim dTotal12
	Dim dTotal13
	Dim dTotal14
	Dim dSMG15
	Dim dSMG30
	Dim dISRFromTable
	Dim dRateFromTable
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bEmpty = True
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "(Companies.", "(Employees.")
	sCondition = Replace(sCondition, "(EmployeeTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(PositionTypes.", "(Employees.")
	sCondition = Replace(sCondition, "(Journeys.", "(Employees.")
	sCondition = Replace(sCondition, "(Shifts.", "(Employees.")
	sCondition = Replace(sCondition, "(Levels.", "(Employees.")
	sCondition = Replace(sCondition, "(PaymentCenters.", "(Employees.")
	sCondition = Replace(sCondition, "(Positions.", "(Jobs.")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN)
	If Not FolderExists(Server.MapPath(sFilePath), sErrorDescription) Then lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFileName = sFilePath & "\Rep_"
	sFilePath = sFilePath & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
	lErrorNumber = CreateFolder(Server.MapPath(sFilePath), sErrorDescription)
	sFilePath = Server.MapPath(sFilePath)
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		Response.Flush()
		sTempFileName = Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt")
		sExcluded = "-1"

		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, InvertedTax, InvertedRate From TaxInvertions Where (StartDate<=" & oRequest("YearID").Item & "0101) And (EndDate>=" & oRequest("YearID").Item & "1231) And (PeriodID=8)", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			asTaxes = ""
			Do While Not oRecordset.EOF
				asTaxes = asTaxes & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("InvertedTax").Value) & "," & CStr(oRecordset.Fields("InvertedRate").Value) & "" & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			asTaxes = asTaxes & "0,1E+69,0,1"
		End If
		asTaxes = Split(asTaxes, ";")
		For iIndex = 0 To UBound(asTaxes)
			asTaxes(iIndex) = Split(asTaxes(iIndex), ",")
			For jIndex = 0 To UBound(asTaxes(iIndex))
				asTaxes(iIndex)(jIndex) = CDbl(asTaxes(iIndex)(jIndex))
			Next
		Next

		lCurrentID = -1
		sErrorDescription = "No se pudieron obtener los registros de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, bTaxAdjustment, Areas.AreaShortName, Zones.ZoneCode, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue From Payroll_" & oRequest("YearID").Item & ", Concepts, Employees, EmployeesForTaxAdjustment, Jobs, Areas, Zones, ZoneTypes, CurrenciesHistoryList Where (Payroll_" & oRequest("YearID").Item & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesForTaxAdjustment.EmployeeID) And (EmployeesForTaxAdjustment.PayrollYear=" & oRequest("YearID").Item & ") And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (ZoneTypes.ZoneTypeID=CurrenciesHistoryList.CurrencyID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (Employees.EmployeeID Not In (" & sExcluded & ")) And (Concepts.ConceptID>0) And ((Concepts.IsDeduction=0) Or (Concepts.ConceptID In (30, 52, 55, 71, 72, 110))) And (CurrenciesHistoryList.CurrencyDate=" & oRequest("YearID").Item & "1231) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID>0) And (Areas.EndDate=30000000) And (Zones.EndDate=30000000) " & sCondition & " Group By Concepts.ConceptID, Concepts.ConceptShortName, Concepts.ConceptName, Concepts.IsDeduction, Concepts.TaxAmount, Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, bTaxAdjustment, Areas.AreaShortName, Zones.ZoneCode, ZoneTypes.ZoneTypeName, CurrenciesHistoryList.CurrencyValue Order By Employees.EmployeeID, Concepts.IsDeduction, Concepts.ConceptShortName", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID > -1 Then
							dTotal2 = dTotal1 - dTotal2
							dTotal3 = dTotal2 - dTotal3
							dTotal4 = dTotal4 - dSMG30
							dTotal5 = dTotal3 + dTotal4
							For iIndex = 0 To UBound(asTaxes)
								If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
									dISRFromTable = asTaxes(iIndex)(2)
									dRateFromTable = asTaxes(iIndex)(3)
									Exit For
								End If
							Next
							dTotal6 = dTotal5 - dISRFromTable
							dTotal7 = dTotal6 / dRateFromTable
							dTotal8 = dTotal7 - dTotal2
							dTotal9 = dTotal8 + dSMG30
							dTotal10 = dTotal8 - dTotal4
							dTotal11 = dTotal1 + dTotal8
							dTotal12 = dTotal12 + dSMG15 + dSMG30
							dTotal13 = dTotal11 + dTotal12
							dTotal14 = dTotal14 + dTotal10
							sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
							sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
							sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
							sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
							For iIndex = 0 To 200
								sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & " />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & "_SN />", "0.00")
								sRowContents = Replace(sRowContents, "<CONCEPT_TAX_" & iIndex & " />", "0.00")
							Next
							lErrorNumber = AppendTextToFile(sTempFileName, sRowContents, sErrorDescription)
						End If
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						sRowContents = ""
						dSMG15 = 0
						dSMG30 = 0
						dTotal1 = 0
						dTotal2 = 0
						dTotal3 = 0
						dTotal4 = 0
						dTotal5 = 0
						dTotal6 = 0
						dTotal7 = 0
						dTotal8 = 0
						dTotal9 = 0
						dTotal10 = 0
						dTotal11 = 0
						dTotal12 = 0
						dTotal13 = 0
						dTotal14 = 0
						sRowContents = "01|12" 'Campo 1 y Campo 2
						sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("RFC").Value) 'Campo 3
						sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("CURP").Value) 'Campo 4
						sRowContents = sRowContents & "|" & Left(CStr(oRecordset.Fields("EmployeeLastName").Value), 43) 'Campo 5
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "|" & Left(CStr(oRecordset.Fields("EmployeeLastName2").Value), 43) 'Campo 6
						Else
							sRowContents = sRowContents & "| " 'Campo 6
						End If
						sRowContents = sRowContents & "|" & Left(CStr(oRecordset.Fields("EmployeeName").Value), 43) 'Campo 7
						sRowContents = sRowContents & "|" & Right(("00" & CStr(oRecordset.Fields("EconomicZoneID").Value)), Len("00")) 'Campo 8
						sRowContents = sRowContents & "|" & Replace(CStr(oRecordset.Fields("bTaxAdjustment").Value), "0", "2") 'Campo 9
						sRowContents = sRowContents & "|1" 'Campo 10
						sRowContents = sRowContents & "|1" 'Campo 11
						sRowContents = sRowContents & "|0.00" 'Campo 12
						If InStr(1, sSyndicateIDs, "," & CStr(oRecordset.Fields("EmployeeID").Value) & ",", vbBinaryCompare) > 0 Then
							sRowContents = sRowContents & "|1" 'Campo 13
						Else
							sRowContents = sRowContents & "|2" 'Campo 13
						End If
						sRowContents = sRowContents & "|0" 'Campo 14
						sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("ZoneCode").Value) 'Campo 15
						sRowContents = sRowContents & "||||||||||" 'Campo 16 - Campo 25
						sRowContents = sRowContents & "|0.00" 'Campo 26
						sRowContents = sRowContents & "|0" 'Campo 27
						sRowContents = sRowContents & "|0.00" 'Campo 28
						sRowContents = sRowContents & "|0.00" 'Campo 29
						sRowContents = sRowContents & "|<CONCEPT_120_SN />" 'Campo 30
						sRowContents = sRowContents & "|2" 'Campo 31
						sRowContents = sRowContents & "|1" 'Campo 32
						sRowContents = sRowContents & "|<TOTAL_1 />" 'Campo 33
						sRowContents = sRowContents & "|0.00" 'Campo 34
						sRowContents = sRowContents & "|0.00" 'Campo 35
						sRowContents = sRowContents & "|0.00" 'Campo 36
						If (CInt(oRequest("YearID").Item) Mod 4) = 0 Then
							sRowContents = sRowContents & "|366" 'Campo 37
						Else
							sRowContents = sRowContents & "|365" 'Campo 37
						End If
						sRowContents = sRowContents & "|<TOTAL_3 />" 'Campo 38
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 39
						sRowContents = sRowContents & "|0.00" 'Campo 40
						sRowContents = sRowContents & "|<TOTAL_1 />" 'Campo 41
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 42
						sRowContents = sRowContents & "|0.00" 'Campo 43
						sRowContents = sRowContents & "|" & CalculateAgeFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), CLng(oRequest("YearID").Item & "1231")) 'Campo 44
						sRowContents = sRowContents & "|<TOTAL_3 />" 'Campo 45
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 46
						sRowContents = sRowContents & "|<TOTAL_1 />" 'Campo 47
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 48
						sRowContents = sRowContents & "|<TOTAL_1 />" 'Campo 49
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 50
						sRowContents = sRowContents & "|<CONCEPT_1 />" 'Campo 51
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 52
						sRowContents = sRowContents & "|2" 'Campo 53
						sRowContents = sRowContents & "|0.00" 'Campo 54
						sRowContents = sRowContents & "|0.00" 'Campo 55
						sRowContents = sRowContents & "|0.00" 'Campo 56
						sRowContents = sRowContents & "|0.00" 'Campo 57
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 58
						sRowContents = sRowContents & "|<TOTAL_3 />" 'Campo 59
						sRowContents = sRowContents & "|<TOTAL_4 />" 'Campo 60
						sRowContents = sRowContents & "|<TOTAL_4 />" 'Campo 61
						sRowContents = sRowContents & "|0.00" 'Campo 62
						sRowContents = sRowContents & "|<CONCEPT_TAX_139 />" 'Campo 63
						sRowContents = sRowContents & "|<CONCEPT_TAX_9 />" 'Campo 64
						sRowContents = sRowContents & "|<CONCEPT_TAX_148 />" 'Campo 65
						sRowContents = sRowContents & "|<CONCEPT_TAX_21 />" 'Campo 66
						sRowContents = sRowContents & "|<CONCEPT_TAX_20 />" 'Campo 67
						sRowContents = sRowContents & "|<CONCEPT_TAX_17 />" 'Campo 68
						sRowContents = sRowContents & "|<CONCEPT_TAX_16 />" 'Campo 69
						sRowContents = sRowContents & "|0.00" 'Campo 70
						sRowContents = sRowContents & "|0.00" 'Campo 71
						sRowContents = sRowContents & "|0.00" 'Campo 72
						sRowContents = sRowContents & "|0.00" 'Campo 73
						sRowContents = sRowContents & "|<CONCEPT_77 />" 'Campo 74
						sRowContents = sRowContents & "|0.00" 'Campo 75
						sRowContents = sRowContents & "|0.00" 'Campo 76
						sRowContents = sRowContents & "|0.00" 'Campo 77
						sRowContents = sRowContents & "|<CONCEPT_125 />" 'Campo 78
						sRowContents = sRowContents & "|0.00" 'Campo 79
						sRowContents = sRowContents & "|0.00" 'Campo 80
						sRowContents = sRowContents & "|<CONCEPT_45 />" 'Campo 81
						sRowContents = sRowContents & "|0.00" 'Campo 82
						sRowContents = sRowContents & "|0.00" 'Campo 83
						sRowContents = sRowContents & "|<CONCEPT_41 />" 'Campo 84
						sRowContents = sRowContents & "|0.00" 'Campo 85
						sRowContents = sRowContents & "|<CONCEPT_65 />" 'Campo 86
						sRowContents = sRowContents & "|0.00" 'Campo 87
						sRowContents = sRowContents & "|<CONCEPT_84 />" 'Campo 88
						sRowContents = sRowContents & "|0.00" 'Campo 89
						sRowContents = sRowContents & "|0.00" 'Campo 90
						sRowContents = sRowContents & "|0.00" 'Campo 91
						sRowContents = sRowContents & "|0.00" 'Campo 92
						sRowContents = sRowContents & "|0.00" 'Campo 93
						sRowContents = sRowContents & "|0.00" 'Campo 94
						sRowContents = sRowContents & "|0.00" 'Campo 95
						sRowContents = sRowContents & "|<CONCEPT_15 />" 'Campo 96
						sRowContents = sRowContents & "|0.00" 'Campo 97
						sRowContents = sRowContents & "|<CONCEPT_37 />" 'Campo 98
						sRowContents = sRowContents & "|0.00" 'Campo 99
						sRowContents = sRowContents & "|<CONCEPT_24 />" 'Campo 100
						sRowContents = sRowContents & "|0.00" 'Campo 101
						sRowContents = sRowContents & "|<CONCEPT_36 />" 'Campo 102
						sRowContents = sRowContents & "|0.00" 'Campo 103
						sRowContents = sRowContents & "|<CONCEPT_56 />" 'Campo 104
						sRowContents = sRowContents & "|0.00" 'Campo 105
						sRowContents = sRowContents & "|0.00" 'Campo 106
						sRowContents = sRowContents & "|0.00" 'Campo 107
						sRowContents = sRowContents & "|<CONCEPT_22 />" 'Campo 108
						sRowContents = sRowContents & "|0.00" 'Campo 109
						sRowContents = sRowContents & "|0.00" 'Campo 110
						sRowContents = sRowContents & "|0.00" 'Campo 111
						sRowContents = sRowContents & "|0.00" 'Campo 112
						sRowContents = sRowContents & "|0.00" 'Campo 113
						sRowContents = sRowContents & "|<TOTAL_2 />" 'Campo 114
						sRowContents = sRowContents & "|<TOTAL_3 />" 'Campo 115
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 116
						sRowContents = sRowContents & "|0.00" 'Campo 117
						sRowContents = sRowContents & "|0.00" 'Campo 118
						sRowContents = sRowContents & "|0.00" 'Campo 119
						sRowContents = sRowContents & "|<CONCEPT_48 />" 'Campo 120
						sRowContents = sRowContents & "|<CONCEPT_49 />" 'Campo 121
						sRowContents = sRowContents & "|0.00" 'Campo 122
						sRowContents = sRowContents & "|0.00" 'Campo 123
						sRowContents = sRowContents & "|<CONCEPT_1 />" 'Campo 124
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 125
						sRowContents = sRowContents & "|0.00" 'Campo 126
						sRowContents = sRowContents & "|0.00" 'Campo 127
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 128
						sRowContents = sRowContents & "|0.00" 'Campo 129
						sRowContents = sRowContents & "|0.00" 'Campo 130
						sRowContents = sRowContents & "|<CONCEPT_55 />" 'Campo 131
						sRowContents = sRowContents & "|0.00" 'Campo 132
						sRowContents = sRowContents & "|0.00" 'Campo 133
						sRowContents = sRowContents & "|0.00" 'Campo 134
						sRowContents = sRowContents & "|0.00" 'Campo 135
						sRowContents = sRowContents & "|0.00" 'Campo 136
						sRowContents = sRowContents & "|0.00" 'Campo 137
						sRowContents = sRowContents & "|0.00" 'Campo 138
						sRowContents = sRowContents & "|0.00" 'Campo 139
						sRowContents = sRowContents & "|0.00" 'Campo 140
						sRowContents = sRowContents & "|0.00" 'Campo 141
						sRowContents = sRowContents & "|0.00" 'Campo 142
						sRowContents = sRowContents & "|0.00" 'Campo 143
						sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("EmployeeNumber").Value) 'Campo 144
						sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("AreaShortName").Value) 'Campo 145
						sRowContents = sRowContents & "|1" 'Campo 146
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 20 '18 Prima de vacaciones exenta
'							sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />", 1, -1, vbBinaryCompare)
'							dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
							dSMG15 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 15
							If dSMG15 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG15 = CDbl(oRecordset.Fields("TotalAmount").Value)
'						Case 21 '18 Prima de vacaciones gravable
'							sRowContents = Replace(sRowContents, "<CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><CONCEPTS />")
'							dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 30 '26. Aguinaldo
							sRowContents = Replace(sRowContents, "<CONCEPT_30 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
							dSMG30 = CDbl(oRecordset.Fields("CurrencyValue").Value) * 30
							If dSMG30 > CDbl(oRecordset.Fields("TotalAmount").Value) Then dSMG30 = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 44, 94 '41 Premio antigüedad 25 y 30 años (mes de sueldo), C3 Premios, estimulos y recompensas (recompensa del sistema de evaluación del desempeño)
							sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><PIRAMID_CONCEPTS />")
							dTotal4 = dTotal4 + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 52, 71, 72 '50 Faltas, 70 Retardos, 71 Deducción por cobro de sueldos indebidos
							sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD WIDTH=""100%""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</B> " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "&nbsp;&nbsp;&nbsp;</FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</NOBR></FONT></TD></TR><DEDUCTION_CONCEPTS />")
							dTotal2 = dTotal2 + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 55 '53 Impuesto sobre producto de trabajo (ISR)
							sRowContents = Replace(sRowContents, "<CONCEPT_55 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
							dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 110 'IS ISR patronal del Seguro de Separación Individualizado
							sRowContents = Replace(sRowContents, "<CONCEPT_110 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							dTotal3 = dTotal3 + CDbl(oRecordset.Fields("TotalAmount").Value)
							dTotal14 = dTotal14 + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case Else
							If (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) > 0) Then
								sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal1 = dTotal1 + CDbl(oRecordset.Fields("TotalAmount").Value)
							ElseIf (CInt(oRecordset.Fields("IsDeduction").Value) = 0) And (CDbl(oRecordset.Fields("TaxAmount").Value) = 0) Then
								sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
								dTotal12 = dTotal12 + CDbl(oRecordset.Fields("TotalAmount").Value)
							End If
					End Select
					sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & "_SN />", "1")
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				dTotal2 = dTotal1 - dTotal2
				dTotal3 = dTotal2 - dTotal3
				dTotal4 = dTotal4 - dSMG30
				dTotal5 = dTotal3 + dTotal4
				For iIndex = 0 To UBound(asTaxes)
					If (asTaxes(iIndex)(0) <= dTotal5) And (asTaxes(iIndex)(1) >= dTotal5) Then
						dISRFromTable = asTaxes(iIndex)(2)
						dRateFromTable = asTaxes(iIndex)(3)
						Exit For
					End If
				Next
				dTotal6 = dTotal5 - dISRFromTable
				dTotal7 = dTotal6 / dRateFromTable
				dTotal8 = dTotal7 - dTotal2
				dTotal9 = dTotal8 + dSMG30
				dTotal10 = dTotal8 - dTotal4
				dTotal11 = dTotal1 + dTotal8
				dTotal12 = dTotal12 + dSMG15 + dSMG30
				dTotal13 = dTotal11 + dTotal12
				dTotal14 = dTotal14 + dTotal10
				sRowContents = Replace(sRowContents, "<SMG_15 />", FormatNumber(dSMG15, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<SMG_30 />", FormatNumber(dSMG30, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<ISR_FROM_TABLE />", FormatNumber(dISRFromTable, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<RATE_FROM_TABLE />", FormatNumber(dRateFromTable, 6, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_1 />", FormatNumber(dTotal1, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_2 />", FormatNumber(dTotal2, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_3 />", FormatNumber(dTotal3, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_4 />", FormatNumber(dTotal4, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_5 />", FormatNumber(dTotal5, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_6 />", FormatNumber(dTotal6, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_7 />", FormatNumber(dTotal7, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_8 />", FormatNumber(dTotal8, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_9 />", FormatNumber(dTotal9, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_10 />", FormatNumber(dTotal10, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_11 />", FormatNumber(dTotal11, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_12 />", FormatNumber(dTotal12, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_13 />", FormatNumber(dTotal13, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<TOTAL_14 />", FormatNumber(dTotal14, 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_55 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPT_110 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPTS />", "")
				sRowContents = Replace(sRowContents, "<DEDUCTION_CONCEPTS />", "")
				sRowContents = Replace(sRowContents, "<PIRAMID_CONCEPTS />", "")
				sRowContents = Replace(sRowContents, "<EXEMPT_CONCEPTS />", "")
				For iIndex = 0 To 200
					sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & " />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_" & iIndex & "_SN />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_TAX_" & iIndex & " />", "0.00")
				Next
				lErrorNumber = AppendTextToFile(sTempFileName, sRowContents, sErrorDescription)
				oRecordset.Close
			End If
		End If

		If Not bEmpty Then
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudo guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1157 = lErrorNumber
	Err.Clear
End Function
%>